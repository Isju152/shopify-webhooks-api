from flask import Flask, request, jsonify
import hmac
import hashlib
import base64
from datetime import datetime, timedelta
import os
from dotenv import load_dotenv
import requests
import logging
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

load_dotenv()

app = Flask(__name__)

# Configurar logging SIMPLE
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

logger = logging.getLogger(__name__)

# Variables de Shopify
SHOPIFY_WEBHOOK_SECRET = os.getenv('SHOPIFY_WEBHOOK_SECRET')
SHOPIFY_SHOP_URL = os.getenv('SHOPIFY_SHOP_URL')
SHOPIFY_ACCESS_TOKEN = os.getenv('SHOPIFY_ACCESS_TOKEN')

# Variables de Email
EMAIL_USER = os.getenv('EMAIL_USER')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')
EMAIL_RECIPIENT = os.getenv('EMAIL_RECIPIENT')


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


def enviar_email(asunto, cuerpo):
    """Envía un email con la notificación"""
    try:
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(EMAIL_USER, EMAIL_PASSWORD)
        
        msg = MIMEMultipart()
        msg['From'] = EMAIL_USER
        msg['To'] = EMAIL_RECIPIENT
        msg['Subject'] = asunto
        
        msg.attach(MIMEText(cuerpo, 'html'))
        
        server.send_message(msg)
        server.quit()
        
        logger.info(f"✉️ Email enviado: {asunto}")
        return True
    
    except Exception as e:
        logger.error(f"❌ Error enviando email: {e}")
        return False


@app.route('/webhooks/orders/create', methods=['POST'])
def order_created():
    """Se ejecuta cuando se crea una nueva orden"""
    signature = request.headers.get('X-Shopify-Hmac-SHA256', '')
    
    if not verificar_webhook(request.get_data(), signature):
        logger.warning("⚠️ Webhook rechazado (firma inválida)")
        return 'Unauthorized', 401
    
    orden = request.get_json()
    order_id = orden['id']
    order_number = orden['order_number']
    customer_name = orden['customer']['first_name'] if orden.get('customer') else 'Cliente'
    customer_email = orden['customer']['email'] if orden.get('customer') else 'Sin email'
    total = orden['total_price']
    financial_status = orden['financial_status']
    
    # LOGS DETALLADOS
    logger.info("="*60)
    logger.info(f"✅ NUEVA ORDEN CREADA")
    logger.info(f"   Número: #{order_number}")
    logger.info(f"   Cliente: {customer_name}")
    logger.info(f"   Email: {customer_email}")
    logger.info(f"   Total: ${total}")
    logger.info(f"   Estado de pago: {financial_status}")
    logger.info("="*60)
    
    # ENVIAR EMAIL
    asunto = f"🎉 Nueva Orden #{order_number}"
    cuerpo = f"""
    <html>
        <body style="font-family: Arial; font-size: 14px; color: #333;">
            <h2 style="color: #2ecc71;">¡Nueva Orden Recibida!</h2>
            
            <table style="width: 100%; border-collapse: collapse;">
                <tr>
                    <td style="padding: 10px; font-weight: bold;">Número de orden:</td>
                    <td style="padding: 10px;">#{order_number}</td>
                </tr>
                <tr style="background-color: #f5f5f5;">
                    <td style="padding: 10px; font-weight: bold;">Cliente:</td>
                    <td style="padding: 10px;">{customer_name}</td>
                </tr>
                <tr>
                    <td style="padding: 10px; font-weight: bold;">Email:</td>
                    <td style="padding: 10px;">{customer_email}</td>
                </tr>
                <tr style="background-color: #f5f5f5;">
                    <td style="padding: 10px; font-weight: bold;">Total:</td>
                    <td style="padding: 10px; font-size: 16px; color: #2ecc71;"><strong>${total}</strong></td>
                </tr>
                <tr>
                    <td style="padding: 10px; font-weight: bold;">Estado de pago:</td>
                    <td style="padding: 10px;">{financial_status}</td>
                </tr>
            </table>
            
            <hr style="margin: 20px 0;">
            <p><small>Esta es una notificación automática de tu servidor de webhooks Amarela.</small></p>
        </body>
    </html>
    """
    
    enviar_email(asunto, cuerpo)
    
    return jsonify({'status': 'ok'}), 200


@app.route('/webhooks/orders/updated', methods=['POST'])
def order_updated():
    """Se ejecuta cuando una orden se actualiza"""
    signature = request.headers.get('X-Shopify-Hmac-SHA256', '')
    
    if not verificar_webhook(request.get_data(), signature):
        logger.warning("⚠️ Webhook rechazado (firma inválida)")
        return 'Unauthorized', 401
    
    orden = request.get_json()
    order_id = orden['id']
    order_number = orden['order_number']
    financial_status = orden['financial_status']
    fulfillment_status = orden['fulfillment_status']
    customer_email = orden['customer']['email'] if orden.get('customer') else 'Sin email'
    total = orden['total_price']
    
    # LOGS DETALLADOS
    logger.info("="*60)
    logger.info(f"🔄 ORDEN ACTUALIZADA")
    logger.info(f"   Número: #{order_number}")
    logger.info(f"   Email: {customer_email}")
    logger.info(f"   Total: ${total}")
    logger.info(f"   Estado de pago: {financial_status}")
    logger.info(f"   Estado de envío: {fulfillment_status}")
    logger.info("="*60)
    
    # DETECCIÓN 1: Devolución
    if financial_status == 'refunded':
        logger.info(f"   ⚠️ DEVOLUCIÓN DETECTADA")
        
        asunto = f"⚠️ Devolución detectada - Orden #{order_number}"
        cuerpo = f"""
        <html>
            <body style="font-family: Arial; font-size: 14px; color: #333;">
                <h2 style="color: #e74c3c;">¡Devolución Detectada!</h2>
                
                <p><strong>Número de orden:</strong> #{order_number}</p>
                <p><strong>Cliente:</strong> {customer_email}</p>
                <p><strong>Total:</strong> ${total}</p>
                <p style="color: red;"><strong>⚠️ La orden ha sido reembolsada</strong></p>
                
                <p>Por favor, revisa el estado en Shopify y en Dropi.</p>
                
                <hr>
                <p><small>Esta es una notificación automática de tu servidor de webhooks Amarela.</small></p>
            </body>
        </html>
        """
        
        enviar_email(asunto, cuerpo)
    
    # DETECCIÓN 2: Pago completado
    if financial_status == 'paid':
        logger.info(f"   ✅ PAGO RECIBIDO (Pago Contra Entrega confirmado)")
        
        asunto = f"✅ Pago confirmado - Orden #{order_number}"
        cuerpo = f"""
        <html>
            <body style="font-family: Arial; font-size: 14px; color: #333;">
                <h2 style="color: #2ecc71;">¡Pago Confirmado!</h2>
                
                <p><strong>Número de orden:</strong> #{order_number}</p>
                <p><strong>Cliente:</strong> {customer_email}</p>
                <p><strong>Total pagado:</strong> ${total}</p>
                <p style="color: green;"><strong>✅ El pago ha sido procesado correctamente</strong></p>
                
                <p>La orden está lista para ser procesada.</p>
                
                <hr>
                <p><small>Esta es una notificación automática de tu servidor de webhooks Amarela.</small></p>
            </body>
        </html>
        """
        
        enviar_email(asunto, cuerpo)
    
    # DETECCIÓN 3: Envío completado
    if fulfillment_status == 'fulfilled':
        logger.info(f"   🚚 ENVÍO COMPLETADO")
        
        asunto = f"🚚 Orden enviada - #{order_number}"
        cuerpo = f"""
        <html>
            <body style="font-family: Arial; font-size: 14px; color: #333;">
                <h2 style="color: #3498db;">¡Orden Enviada!</h2>
                
                <p><strong>Número de orden:</strong> #{order_number}</p>
                <p><strong>Cliente:</strong> {customer_email}</p>
                <p style="color: blue;"><strong>🚚 Tu orden ha sido enviada</strong></p>
                
                <p>Verifica tu email para el tracking y más detalles.</p>
                
                <hr>
                <p><small>Esta es una notificación automática de tu servidor de webhooks Amarela.</small></p>
            </body>
        </html>
        """
        
        enviar_email(asunto, cuerpo)
    
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
    
    asunto = f"💰 Reembolso procesado - Orden {order_id}"
    cuerpo = f"""
    <html>
        <body style="font-family: Arial; font-size: 14px; color: #333;">
            <h2 style="color: #f39c12;">Reembolso Procesado</h2>
            
            <p><strong>Orden:</strong> {order_id}</p>
            <p><strong>Monto reembolsado:</strong> ${monto}</p>
            
            <p>El reembolso ha sido procesado correctamente.</p>
            
            <hr>
            <p><small>Esta es una notificación automática de tu servidor de webhooks Amarela.</small></p>
        </body>
    </html>
    """
    
    enviar_email(asunto, cuerpo)
    
    return jsonify({'status': 'ok'}), 200


@app.route('/check-stale-orders', methods=['GET'])
def check_stale_orders():
    """Endpoint que chequea órdenes sin actualizar >24h"""
    
    logger.info("🔍 Iniciando chequeo de órdenes estancadas (>24h)...")
    
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
            
            asunto = f"⚠️ {len(ordenes)} órdenes estancadas"
            cuerpo = f"""
            <html>
                <body style="font-family: Arial; font-size: 14px; color: #333;">
                    <h2 style="color: #e74c3c;">Órdenes Estancadas (>24h)</h2>
                    
                    <p>Se encontraron <strong>{len(ordenes)} órdenes</strong> sin actualizar hace más de 24 horas:</p>
                    
                    <ul>
            """
            
            for orden in ordenes:
                order_number = orden['order_number']
                updated_at = orden['updated_at']
                customer_email = orden['customer']['email'] if orden.get('customer') else 'Sin email'
                
                logger.info(f"   📋 Orden #{order_number}")
                logger.info(f"      - Cliente: {customer_email}")
                logger.info(f"      - Última actualización: {updated_at}")
                
                cuerpo += f"""
                    <li>
                        <strong>Orden #{order_number}</strong> - {customer_email}<br>
                        Última actualización: {updated_at}
                    </li>
                """
            
            cuerpo += """
                    </ul>
                    
                    <p style="color: red;"><strong>⚠️ Por favor, revisa estas órdenes en Shopify y Dropi.</strong></p>
                    
                    <hr>
                    <p><small>Esta es una notificación automática de tu servidor de webhooks Amarela.</small></p>
                </body>
            </html>
            """
            
            enviar_email(asunto, cuerpo)
        else:
            logger.info("✅ Todas las órdenes actualizadas correctamente")
        
        return jsonify({
            'status': 'ok',
            'stale_orders': len(ordenes)
        }), 200
    
    except Exception as e:
        logger.error(f"❌ Error en chequeo de órdenes: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/', methods=['GET'])
def health():
    """Endpoint de salud"""
    logger.info("✅ Health check - Servidor activo")
    return jsonify({
        'status': 'ok',
        'message': 'Servidor de webhooks Shopify activo',
        'shop': SHOPIFY_SHOP_URL
    }), 200


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    logger.info(f"🚀 Iniciando servidor en puerto {port}")
    app.run(host='0.0.0.0', port=port, debug=False)