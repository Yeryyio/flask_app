from flask import Flask, render_template, request, redirect, url_for
from flask_mail import Mail, Message
import pandas as pd
import os

app = Flask(__name__)

# Configuración de Flask-Mail para Gmail
app.config['MAIL_SERVER'] = 'smtp.gmail.com'
app.config['MAIL_PORT'] = 587
app.config['MAIL_USE_TLS'] = True
app.config['MAIL_USE_SSL'] = False
app.config['MAIL_USERNAME'] = 'sergioriquelme328@gmail.com'
app.config['MAIL_PASSWORD'] = 'teor ivor hfas guvt'  # O usa una contraseña de aplicación si tienes 2FA activado

mail = Mail(app)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    orders = request.json['orders']
    df = pd.DataFrame(orders)
    
    # Guardar el DataFrame en un archivo Excel
    file_path = 'pedidos.xlsx'
    df.to_excel(file_path, index=False)

    # Enviar el archivo por correo
    msg = Message('Nuevos Pedidos de Café',
                  sender='sergioriquelme328@gmail.com',
                  recipients=['acardenas.alvica@gmail.com'])
    msg.body = 'Se adjunta el archivo con los pedidos de café.'
    with app.open_resource(file_path) as fp:
        msg.attach(file_path, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', fp.read())
    
    mail.send(msg)
    
    # Eliminar el archivo después de enviarlo
    os.remove(file_path)
    
    return {'message': 'Pedidos enviados correctamente'}

if __name__ == '__main__':
    app.run(debug=True)

'''
from flask import Flask, render_template, request, jsonify
from flask_mail import Mail, Message
from openpyxl import Workbook
from io import BytesIO

app = Flask(__name__)

# Configuración de Flask-Mail para Gmail
app.config['MAIL_SERVER'] = 'smtp.gmail.com'
app.config['MAIL_PORT'] = 587
app.config['MAIL_USE_TLS'] = True
app.config['MAIL_USE_SSL'] = False
app.config['MAIL_USERNAME'] = 'sergioriquelme328@gmail.com'
app.config['MAIL_PASSWORD'] = 'teor ivor hfas guvt'  # O usa una contraseña de aplicación si tienes 2FA activado

mail = Mail(app)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    orders = request.json
    if orders:
        # Crear el archivo Excel
        wb = Workbook()
        ws = wb.active
        ws.title = "Pedidos de Café"
        
        # Agregar encabezados
        headers = ["Tipo de Café", "Cantidad", "Fecha y Hora", "Ubicación", "Razón"]
        ws.append(headers)
        
        # Agregar datos
        for order in orders:
            ws.append([
                order['coffee_type'],
                order['quantity'],
                order['date_time'],
                order['location'],
                order['reason']
            ])
        
        # Guardar el archivo Excel en memoria
        excel_file = BytesIO()
        wb.save(excel_file)
        excel_file.seek(0)
        
        # Construir el mensaje del correo electrónico
        msg = Message('Nuevo Pedido de Café',
                      sender='sergioriquelme328@gmail.com',
                      recipients=['it.alvica2024@gmail.com'])
        msg.body = 'Adjunto encontrará un archivo Excel con los detalles de los pedidos.'
        
        # Adjuntar el archivo Excel
        msg.attach("pedidos_cafe.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", excel_file.getvalue())
        
        # Enviar el correo
        mail.send(msg)
        
        return jsonify({'status': 'success'}), 200
    else:
        return jsonify({'status': 'error', 'message': 'No orders found'}), 400

if __name__ == '__main__':
    app.run(debug=True)
'''























'''from flask import Flask, render_template, request, redirect, url_for
from flask_mail import Mail, Message
import pandas as pd
from io import BytesIO
import base64

app = Flask(__name__)

# Configuración de Flask-Mail para Gmail
app.config['MAIL_SERVER'] = 'smtp.gmail.com'
app.config['MAIL_PORT'] = 587
app.config['MAIL_USE_TLS'] = True
app.config['MAIL_USE_SSL'] = False
app.config['MAIL_USERNAME'] = 'sergioriquelme328@gmail.com'
app.config['MAIL_PASSWORD'] = 'teor ivor hfas guvt'  # Usa una contraseña de aplicación si tienes 2FA activado

mail = Mail(app)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    orders = request.form.getlist('orders')
    if not orders:
        return redirect(url_for('index'))

    # Crear un DataFrame para los pedidos
    df = pd.DataFrame([order.split(',') for order in orders],
                      columns=['Razón', 'Tipo de Café', 'Cantidad', 'Fecha y Hora', 'Ubicación'])

    # Crear un archivo Excel en memoria
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Pedidos')

    # Enviar el correo
    output.seek(0)
    msg = Message('Nuevo Pedido de Café',
                  sender='sergioriquelme328@gmail.com',
                  recipients=['acardenas.alvica@gmail.com'])
    msg.body = 'Adjunto encontrarás un archivo Excel con los pedidos cargados.'
    msg.attach('pedidos.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', output.read())
    mail.send(msg)

    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
'''










