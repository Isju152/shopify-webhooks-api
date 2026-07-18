from flask import Flask, request, jsonify
import hmac
import hashlib
import base64
from datetime import datetime, timedelta
import os
from dotenv import load_dotenv
import requests
import logging
import sys
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

load_dotenv()

app = Flask(__name__)

# Configurar logging CORRECTAMENTE para Render
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    stream=sys.stdout,
    force=True  # Fuerza la configuración
)

# También configurar stderr
handler = logging.StreamHandler(sys.stderr)
handler.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)

logger = logging.getLogger(__name__)
logger.addHandler(handler)
logger.setLevel(logging.INFO)

# Desactivar buffering
sys.stdout = os.fdopen(sys.stdout.fileno(), 'w', 0)
sys.stderr = os.fdopen(sys.stderr.fileno(), 'w', 0)

app = Flask(__name__)

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


# ===============================
# WEBHOOK 1: Nueva orden
# ===============================
@app.route('/webhooks/orders/create', methods=['POST'])
def order_created():
    """Se ejecuta cuando se crea una nueva orden"""
    signature = request.headers.get('X-Shopify-Hmac-SHA256', '')
    
    # Verificar que es legítimo
    if not verificar_webhook(request.get_data(), signature):
        print("⚠️ Webhook rechazado (firma inválida)")
        return 'Unauthorized', 401
    
    orden = request.get_json()
    order_id = orden['id']
    order_number = orden['order_number']
    customer_email = orden['customer']['email'] if orden.get('customer') else 'Sin email'
    total = orden['total_price']
    
    print(f"\n✅ NUEVA ORDEN #{order_number}")
    print(f"   Cliente: {customer_email}")
    print(f"   Total: ${total}")
    print(f"   Pago: {orden['financial_status']}")
    
    # AQUÍ: Puedes agregar tu lógica
    # Por ejemplo: enviar a proveedor, crear ticket, etc
    
    return jsonify({'status': 'ok'}), 200


# ===============================
# WEBHOOK 2: Orden actualizada
# ===============================
@app.route('/webhooks/orders/updated', methods=['POST'])
def order_updated():
    """Se ejecuta cuando una orden se actualiza"""
    signature = request.headers.get('X-Shopify-Hmac-SHA256', '')
    
    if not verificar_webhook(request.get_data(), signature):
        print("⚠️ Webhook rechazado (firma inválida)")
        return 'Unauthorized', 401
    
    orden = request.get_json()
    order_id = orden['id']
    order_number = orden['order_number']
    financial_status = orden['financial_status']  # authorized, paid, refunded, etc
    fulfillment_status = orden['fulfillment_status']  # fulfilled, partial, unshipped, etc
    
    print(f"\n🔄 ORDEN ACTUALIZADA #{order_number}")
    print(f"   Estado de pago: {financial_status}")
    print(f"   Estado de envío: {fulfillment_status}")
    
    # DETECCIÓN 1: Devolución
    if financial_status == 'refunded':
        print(f"   ⚠️ DEVOLUCIÓN DETECTADA")
        # Aquí: Crear caso, notificar, etc
    
    # DETECCIÓN 2: Pago completado
    if financial_status == 'paid':
        print(f"   ✅ PAGO RECIBIDO (pago contra entrega)")
        # Aquí: Procesar pago, actualizar estado, etc
    
    # DETECCIÓN 3: Envío completado
    if fulfillment_status == 'fulfilled':
        print(f"   🚚 ENVÍO COMPLETADO")
        # Aquí: Notificar cliente, generar reporte, etc
    
    return jsonify({'status': 'ok'}), 200


# ===============================
# WEBHOOK 3: Devolución/Reembolso
# ===============================
@app.route('/webhooks/refunds/create', methods=['POST'])
def refund_created():
    """Se ejecuta cuando se crea un reembolso"""
    signature = request.headers.get('X-Shopify-Hmac-SHA256', '')
    
    if not verificar_webhook(request.get_data(), signature):
        return 'Unauthorized', 401
    
    refund = request.get_json()
    order_id = refund['order_id']
    
    print(f"\n💰 REEMBOLSO CREADO")
    print(f"   Orden: {order_id}")
    print(f"   Monto: ${refund['transactions'][0]['amount']}")
    
    # Aquí: Crear RMA, alertar proveedor, etc
    
    return jsonify({'status': 'ok'}), 200


# ===============================
# ENDPOINT: Chequear órdenes estancadas
# ===============================
@app.route('/check-stale-orders', methods=['GET'])
def check_stale_orders():
    """
    Endpoint que chequea órdenes sin actualizar >24h
    Puedes llamarlo manualmente o con un CRON job
    """
    
    print("\n🔍 Chequeo de órdenes estancadas (>24h)...")
    
    headers = {
        "X-Shopify-Access-Token": SHOPIFY_ACCESS_TOKEN,
        "Content-Type": "application/json"
    }
    
    # Calcular fecha límite (hace 24 horas)
    limit_time = datetime.now() - timedelta(hours=24)
    limit_iso = limit_time.isoformat()
    
    # Pedir órdenes sin actualizar
    url = f"https://{SHOPIFY_SHOP_URL}/admin/api/2024-01/orders.json"
    params = {
        "status": "any",
        "updated_at_max": limit_iso,  # Solo órdenes ANTES de hace 24h
        "limit": 250
    }
    
    try:
        response = requests.get(url, headers=headers, params=params)
        ordenes = response.json().get('orders', [])
        
        if ordenes:
            print(f"⚠️ Encontradas {len(ordenes)} órdenes sin actualizar >24h:")
            
            for orden in ordenes:
                order_id = orden['id']
                order_number = orden['order_number']
                updated_at = orden['updated_at']
                customer_email = orden['customer']['email'] if orden.get('customer') else 'Sin email'
                
                print(f"\n   Orden #{order_number}")
                print(f"   - Cliente: {customer_email}")
                print(f"   - Última actualización: {updated_at}")
                print(f"   - Estado: {orden['financial_status']}")
                
                # AQUÍ: Tu lógica
                # crear_caso_quality(order_id)
                # send_slack_alert(f"Orden #{order_number} estancada")
                # send_email(f"Orden estancada: #{order_number}")
        else:
            print("✅ Todas las órdenes actualizadas correctamente")
        
        return jsonify({
            'status': 'ok',
            'stale_orders': len(ordenes)
        }), 200
    
    except Exception as e:
        print(f"❌ Error: {e}")
        return jsonify({'error': str(e)}), 500


# ===============================
# PÁGINA DE SALUD
# ===============================
@app.route('/', methods=['GET'])
def health():
    """Simple endpoint para verificar que el servidor está funcionando"""
    return jsonify({
        'status': 'ok',
        'message': 'Servidor de webhooks Shopify activo',
        'shop': SHOPIFY_SHOP_URL
    }), 200


# ===============================
# EJECUTAR SERVIDOR
# ===============================
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)