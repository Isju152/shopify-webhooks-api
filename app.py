from flask import Flask, request, jsonify, render_template_string, send_file
import hmac
import hashlib
import base64
from datetime import datetime, timedelta
import os
from dotenv import load_dotenv
import requests
import logging
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from io import BytesIO

load_dotenv()

app = Flask(__name__)

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

logger = logging.getLogger(__name__)

# Variables de Shopify
SHOPIFY_WEBHOOK_SECRET = os.getenv('SHOPIFY_WEBHOOK_SECRET')
SHOPIFY_SHOP_URL = os.getenv('SHOPIFY_SHOP_URL')
SHOPIFY_ACCESS_TOKEN = os.getenv('SHOPIFY_ACCESS_TOKEN')


def verificar_webhook(request_body, signature):
    """Verifica que el webhook sea legítimo de Shopify"""
    hash_calculated = base64.b64encode(
        hmac.new(
            SHOPIFY_WEBHOOK_SECRET.encode('utf-8'),
            request_body,
            hashlib.sha256
        ).digest()
    ).decode()
    
    return hmac.compare_digest(hash_calculated, signature)


def generar_reporte_excel_ordenes_estancadas():
    """Genera un Excel con órdenes sin actualizar >24h"""
    
    logger.info("📊 Generando reporte de órdenes estancadas...")
    
    headers = {
        "X-Shopify-Access-Token": SHOPIFY_ACCESS_TOKEN,
        "Content-Type": "application/json"
    }
    
    limit_time = datetime.now() - timedelta(hours=24)
    limit_iso = limit_time.isoformat()
    
    url = f"https://{SHOPIFY_SHOP_URL}/admin/api/2024-01/orders.json"
    params = {
        "status": "any",
        "updated_at_max": limit_iso,
        "limit": 250
    }
    
    try:
        response = requests.get(url, headers=headers, params=params)
        ordenes = response.json().get('orders', [])
        
        # Crear workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Órdenes Estancadas"
        
        # Estilos
        header_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=12)
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Encabezados
        headers_excel = [
            "Número Orden",
            "Cliente",
            "Teléfono",
            "Email",
            "Total",
            "Estado Pago",
            "Estado Envío",
            "Última Actualización",
            "Horas Sin Actualizar",
            "Tags"
        ]
        
        ws.append(headers_excel)
        
        # Aplicar estilos a encabezados
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border
        
        # Agregar datos
        for orden in ordenes:
            order_number = orden['order_number']
            customer_name = orden['customer']['first_name'] if orden.get('customer') else 'Sin cliente'
            
            # Obtener teléfono
            customer_phone = None
            if orden.get('customer'):
                customer_phone = orden['customer'].get('phone')
            if not customer_phone:
                customer_phone = 'Sin teléfono'
            
            customer_email = orden['customer'].get('email') if orden.get('customer') else 'Sin email'
            total = orden['total_price']
            financial_status = orden['financial_status']
            fulfillment_status = orden['fulfillment_status']
            updated_at = orden['updated_at']
            tags = orden.get('tags', '')
            
            # Calcular horas sin actualizar
            updated_datetime = datetime.fromisoformat(updated_at.replace('Z', '+00:00'))
            horas_sin_actualizar = (datetime.now(updated_datetime.tzinfo) - updated_datetime).total_seconds() / 3600
            
            row = [
                order_number,
                customer_name,
                customer_phone,
                customer_email,
                f"${total}",
                financial_status,
                fulfillment_status,
                updated_at,
                f"{horas_sin_actualizar:.1f}h",
                tags
            ]
            
            ws.append(row)
            
            # Aplicar bordes
            for cell in ws[ws.max_row]:
                cell.border = border
                cell.alignment = Alignment(horizontal='left', vertical='center')
        
        # Ajustar ancho de columnas
        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 20
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 25
        ws.column_dimensions['E'].width = 12
        ws.column_dimensions['F'].width = 15
        ws.column_dimensions['G'].width = 15
        ws.column_dimensions['H'].width = 25
        ws.column_dimensions['I'].width = 18
        ws.column_dimensions['J'].width = 20
        
        # Guardar en memoria
        excel_file = BytesIO()
        wb.save(excel_file)
        excel_file.seek(0)
        
        logger.info(f"✅ Reporte generado con {len(ordenes)} órdenes estancadas")
        
        return excel_file
    
    except Exception as e:
        logger.error(f"❌ Error generando reporte: {e}")
        return None


@app.route('/webhooks/orders/create', methods=['POST'])
def order_created():
    """Se ejecuta cuando se crea una nueva orden"""
    signature = request.headers.get('X-Shopify-Hmac-SHA256', '')
    
    if not verificar_webhook(request.get_data(), signature):
        logger.warning("⚠️ Webhook rechazado (firma inválida)")
        return 'Unauthorized', 401
    
    orden = request.get_json()
    order_number = orden['order_number']
    customer_name = orden['customer']['first_name'] if orden.get('customer') else 'Cliente'
    
    # Obtener teléfono
    customer_phone = None
    if orden.get('customer'):
        customer_phone = orden['customer'].get('phone')
    if not customer_phone:
        customer_phone = 'Sin teléfono'
    
    total = orden['total_price']
    financial_status = orden['financial_status']
    
    logger.info("="*60)
    logger.info(f"✅ NUEVA ORDEN CREADA")
    logger.info(f"   Número: #{order_number}")
    logger.info(f"   Cliente: {customer_name}")
    logger.info(f"   Teléfono: {customer_phone}")
    logger.info(f"   Total: ${total}")
    logger.info(f"   Estado de pago: {financial_status}")
    logger.info("="*60)
    
    return jsonify({'status': 'ok'}), 200


@app.route('/webhooks/orders/updated', methods=['POST'])
def order_updated():
    """Se ejecuta cuando una orden se actualiza"""
    signature = request.headers.get('X-Shopify-Hmac-SHA256', '')
    
    if not verificar_webhook(request.get_data(), signature):
        logger.warning("⚠️ Webhook rechazado (firma inválida)")
        return 'Unauthorized', 401
    
    orden = request.get_json()
    order_number = orden['order_number']
    financial_status = orden['financial_status']
    fulfillment_status = orden['fulfillment_status']
    customer_name = orden['customer']['first_name'] if orden.get('customer') else 'Cliente'
    
    # Obtener teléfono
    customer_phone = None
    if orden.get('customer'):
        customer_phone = orden['customer'].get('phone')
    if not customer_phone:
        customer_phone = 'Sin teléfono'
    
    total = orden['total_price']
    tags = orden.get('tags', '')
    
    logger.info("="*60)
    logger.info(f"🔄 ORDEN ACTUALIZADA")
    logger.info(f"   Número: #{order_number}")
    logger.info(f"   Teléfono: {customer_phone}")
    logger.info(f"   Total: ${total}")
    logger.info(f"   Estado de pago: {financial_status}")
    logger.info(f"   Estado de envío: {fulfillment_status}")
    if tags:
        logger.info(f"   Tags: {tags}")
    logger.info("="*60)
    
    # DETECCIÓN 1: Asignado a mensajero
    if tags and 'asignado a mensajero' in tags.lower():
        logger.info(f"   🚚 ASIGNADO A MENSAJERO - Cliente: {customer_name}")
    
    # DETECCIÓN 2: En ruta de entrega
    if tags and 'en ruta de entrega' in tags.lower():
        logger.info(f"   📍 EN RUTA DE ENTREGA - Cliente: {customer_name}")
    
    # DETECCIÓN 3: Devolución
    if financial_status == 'refunded':
        logger.info(f"   ⚠️ DEVOLUCIÓN DETECTADA - Cliente: {customer_name}")
    
    # DETECCIÓN 4: Pago completado
    if financial_status == 'paid':
        logger.info(f"   ✅ PAGO RECIBIDO - Cliente: {customer_name}")
    
    # DETECCIÓN 5: Envío completado
    if fulfillment_status == 'fulfilled':
        logger.info(f"   🎉 ENVÍO COMPLETADO - Cliente: {customer_name}")
    
    return jsonify({'status': 'ok'}), 200


@app.route('/webhooks/refunds/create', methods=['POST'])
def refund_created():
    """Se ejecuta cuando se crea un reembolso"""
    signature = request.headers.get('X-Shopify-Hmac-SHA256', '')
    
    if not verificar_webhook(request.get_data(), signature):
        return 'Unauthorized', 401
    
    refund = request.get_json()
    order_id = refund['order_id']
    monto = refund['transactions'][0]['amount']
    
    logger.info("="*60)
    logger.info(f"💰 REEMBOLSO CREADO")
    logger.info(f"   Orden: {order_id}")
    logger.info(f"   Monto: ${monto}")
    logger.info("="*60)
    
    return jsonify({'status': 'ok'}), 200


@app.route('/check-stale-orders', methods=['GET'])
def check_stale_orders():
    """Endpoint para chequear órdenes estancadas manualmente"""
    
    logger.info("🔍 Chequeo manual de órdenes estancadas (>24h)...")
    
    headers = {
        "X-Shopify-Access-Token": SHOPIFY_ACCESS_TOKEN,
        "Content-Type": "application/json"
    }
    
    limit_time = datetime.now() - timedelta(hours=24)
    limit_iso = limit_time.isoformat()
    
    url = f"https://{SHOPIFY_SHOP_URL}/admin/api/2024-01/orders.json"
    params = {
        "status": "any",
        "updated_at_max": limit_iso,
        "limit": 250
    }
    
    try:
        response = requests.get(url, headers=headers, params=params)
        ordenes = response.json().get('orders', [])
        
        if ordenes:
            logger.info(f"⚠️ Encontradas {len(ordenes)} órdenes sin actualizar >24h")
            
            for orden in ordenes:
                order_number = orden['order_number']
                updated_at = orden['updated_at']
                customer_name = orden['customer']['first_name'] if orden.get('customer') else 'Sin cliente'
                
                logger.info(f"   📋 Orden #{order_number} - {customer_name} - Actualización: {updated_at}")
        else:
            logger.info("✅ Todas las órdenes actualizadas correctamente")
        
        return jsonify({
            'status': 'ok',
            'stale_orders': len(ordenes)
        }), 200
    
    except Exception as e:
        logger.error(f"❌ Error en chequeo de órdenes: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/descargar-reporte-estancadas', methods=['GET'])
def descargar_reporte_estancadas():
    """Endpoint para descargar reporte de órdenes estancadas"""
    
    logger.info("📥 Descargando reporte de órdenes estancadas...")
    
    excel_file = generar_reporte_excel_ordenes_estancadas()
    
    if excel_file:
        return send_file(
            excel_file,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'ordenes_estancadas_{datetime.now().strftime("%Y-%m-%d_%H-%M-%S")}.xlsx'
        )
    else:
        return jsonify({'error': 'Error al generar reporte'}), 500


@app.route('/dashboard', methods=['GET'])
def dashboard():
    """Página con botones para chequear y descargar reportes"""
    html = """
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Dashboard - Amarela Webhooks</title>
        <style>
            * {
                margin: 0;
                padding: 0;
                box-sizing: border-box;
            }
            
            body {
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                min-height: 100vh;
                display: flex;
                justify-content: center;
                align-items: center;
                padding: 20px;
            }
            
            .container {
                background: white;
                border-radius: 10px;
                box-shadow: 0 10px 40px rgba(0, 0, 0, 0.2);
                padding: 40px;
                max-width: 500px;
                width: 100%;
            }
            
            h1 {
                color: #333;
                margin-bottom: 10px;
                text-align: center;
            }
            
            .subtitle {
                color: #666;
                text-align: center;
                margin-bottom: 40px;
                font-size: 14px;
            }
            
            .status {
                background: #f0f7ff;
                border-left: 4px solid #2ecc71;
                padding: 15px;
                margin-bottom: 30px;
                border-radius: 5px;
            }
            
            .status.success {
                border-left-color: #2ecc71;
                color: #27ae60;
            }
            
            .buttons-group {
                display: flex;
                flex-direction: column;
                gap: 15px;
            }
            
            button {
                width: 100%;
                padding: 15px;
                border: none;
                border-radius: 5px;
                font-size: 16px;
                font-weight: bold;
                cursor: pointer;
                transition: transform 0.2s, box-shadow 0.2s;
            }
            
            .btn-check {
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                color: white;
            }
            
            .btn-report {
                background: linear-gradient(135deg, #27ae60 0%, #229954 100%);
                color: white;
            }
            
            button:hover:not(:disabled) {
                transform: translateY(-2px);
                box-shadow: 0 5px 20px rgba(0, 0, 0, 0.2);
            }
            
            button:disabled {
                opacity: 0.6;
                cursor: not-allowed;
            }
            
            .result {
                margin-top: 30px;
                padding: 20px;
                background: #f8f9fa;
                border-radius: 5px;
                display: none;
            }
            
            .result.show {
                display: block;
            }
            
            .result h3 {
                color: #333;
                margin-bottom: 10px;
            }
            
            .result-content {
                color: #555;
                font-size: 14px;
                line-height: 1.6;
            }
            
            .loading-spinner {
                display: inline-block;
                width: 20px;
                height: 20px;
                border: 3px solid #f3f3f3;
                border-top: 3px solid #fff;
                border-radius: 50%;
                animation: spin 1s linear infinite;
                margin-right: 10px;
            }
            
            @keyframes spin {
                0% { transform: rotate(0deg); }
                100% { transform: rotate(360deg); }
            }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>📊 Dashboard Amarela</h1>
            <p class="subtitle">Sistema de monitoreo de webhooks y reportes</p>
            
            <div class="status success">
                ✅ Servidor activo y funcionando correctamente
            </div>
            
            <div class="buttons-group">
                <button class="btn-check" id="checkButton" onclick="checkStaleOrders()">
                    🔍 Chequear Órdenes Estancadas (>24h)
                </button>
                
                <button class="btn-report" onclick="descargarReporte()">
                    📊 Descargar Reporte en Excel
                </button>
            </div>
            
            <div class="result" id="result">
                <h3>Resultado del Chequeo</h3>
                <div class="result-content" id="resultContent"></div>
            </div>
        </div>
        
        <script>
            async function checkStaleOrders() {
                const button = document.getElementById('checkButton');
                const result = document.getElementById('result');
                const resultContent = document.getElementById('resultContent');
                
                button.disabled = true;
                button.innerHTML = '<span class="loading-spinner"></span>Ejecutando chequeo...';
                result.classList.remove('show');
                
                try {
                    const response = await fetch('/check-stale-orders');
                    const data = await response.json();
                    
                    if (data.stale_orders === 0) {
                        resultContent.innerHTML = `
                            <p style="color: #27ae60; font-weight: bold;">✅ Perfecto</p>
                            <p>No hay órdenes estancadas. Todas las órdenes han sido actualizadas correctamente.</p>
                        `;
                    } else {
                        resultContent.innerHTML = `
                            <p style="color: #e74c3c; font-weight: bold;">⚠️ Órdenes Estancadas</p>
                            <p>Se encontraron <strong>${data.stale_orders}</strong> órdenes sin actualizar en más de 24 horas.</p>
                            <p style="margin-top: 10px; color: #666; font-size: 12px;">Descarga el reporte para ver más detalles.</p>
                        `;
                    }
                    
                    result.classList.add('show');
                } catch (error) {
                    resultContent.innerHTML = `
                        <p style="color: #e74c3c; font-weight: bold;">❌ Error</p>
                        <p>${error.message}</p>
                    `;
                    result.classList.add('show');
                } finally {
                    button.disabled = false;
                    button.innerHTML = '🔍 Chequear Órdenes Estancadas (>24h)';
                }
            }
            
            function descargarReporte() {
                const button = event.target;
                button.disabled = true;
                button.innerHTML = '<span class="loading-spinner"></span>Generando reporte...';
                
                setTimeout(() => {
                    window.location.href = '/descargar-reporte-estancadas';
                    button.disabled = false;
                    button.innerHTML = '📊 Descargar Reporte en Excel';
                }, 1000);
            }
        </script>
    </body>
    </html>
    """
    return render_template_string(html)


@app.route('/', methods=['GET'])
def health():
    """Endpoint de salud"""
    return jsonify({
        'status': 'ok',
        'message': 'Servidor de webhooks Shopify activo',
        'shop': SHOPIFY_SHOP_URL,
        'dashboard': 'https://shopify-webhooks-api.onrender.com/dashboard'
    }), 200


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    logger.info(f"🚀 Iniciando servidor en puerto {port}")
    app.run(host='0.0.0.0', port=port, debug=False)