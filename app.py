"""WhatsApp webhook (Twilio or Meta Cloud API). Run ONE worker: confirmations are kept in memory."""
import os, hmac, hashlib, base64
from xml.sax.saxutils import escape
from flask import Flask, request, abort, Response
from chama_core import Bot, SheetsBackend, last9, split_messages

app = Flask(__name__)
PROVIDER = os.environ.get('PROVIDER', 'twilio')
bot = Bot(SheetsBackend(os.environ['SHEET_ID'], os.environ.get('GOOGLE_CREDS', 'service_account.json')),
          [a for a in os.environ.get('ADMIN_NUMBERS', '').split(',') if a.strip()])
seen = []  # message ids already processed (providers retry; never double-record a payment)
def first_time(mid):
    if mid in seen: return False
    seen.append(mid); del seen[:-500]; return True

@app.get('/health')
def health(): return 'ok'

# ---- Twilio WhatsApp
def twilio_ok():
    token = os.environ.get('TWILIO_AUTH_TOKEN', '')
    url = os.environ.get('PUBLIC_URL', request.url)
    data = url + ''.join(k + v for k, v in sorted(request.form.items()))
    sig = base64.b64encode(hmac.new(token.encode(), data.encode(), hashlib.sha1).digest()).decode()
    return hmac.compare_digest(sig, request.headers.get('X-Twilio-Signature', ''))

# ---- Meta WhatsApp Cloud API
def meta_ok():
    secret = os.environ.get('META_APP_SECRET', '')
    exp = 'sha256=' + hmac.new(secret.encode(), request.get_data(), hashlib.sha256).hexdigest()
    return hmac.compare_digest(exp, request.headers.get('X-Hub-Signature-256', ''))
def meta_send(to, text):
    import requests
    for part in split_messages(text):
        requests.post(f"https://graph.facebook.com/v20.0/{os.environ['META_PHONE_ID']}/messages",
            headers={'Authorization': 'Bearer ' + os.environ['META_TOKEN']},
            json={'messaging_product': 'whatsapp', 'to': to, 'type': 'text', 'text': {'body': part}}, timeout=15)

@app.get('/whatsapp')
def verify():  # Meta webhook verification handshake
    if request.args.get('hub.verify_token') == os.environ.get('META_VERIFY_TOKEN'): return request.args.get('hub.challenge', '')
    abort(403)

@app.post('/whatsapp')
def inbound():
    if PROVIDER == 'twilio':
        if not twilio_ok(): abort(403)
        if not first_time(request.form.get('MessageSid')): return Response('<Response/>', mimetype='text/xml')
        reply = bot.handle(request.form['From'].replace('whatsapp:', ''), request.form.get('Body', ''))
        messages = ''.join(f'<Message>{escape(part)}</Message>' for part in split_messages(reply))
        return Response(f'<Response>{messages}</Response>', mimetype='text/xml')
    if not meta_ok(): abort(403)
    for e in request.json.get('entry', []):
        for ch in e.get('changes', []):
            for msg in ch.get('value', {}).get('messages', []):
                if msg.get('type') == 'text' and first_time(msg['id']):
                    meta_send(msg['from'], bot.handle(msg['from'], msg['text']['body']))
    return 'ok'

if __name__ == '__main__':
    app.run(port=int(os.environ.get('PORT', 5000)))
