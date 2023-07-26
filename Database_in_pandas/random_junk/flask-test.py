from flask import Flask
from flask_mail import Mail, Message

app = Flask(__name__)

#config
app.config['MAIL_SERVER'] = 'smtp.gmail.com'
app.config['MAIL_PORT'] = 465
app.config['MAIL_USERNAME'] = 'jambo5167@gmail.com'
app.config['MAIL_PASSWORD'] = 'Famasf123456'
app.config['MAIL_USE_TLS'] = False
app.config['MAIL_USE_SSL'] = True

mail = Mail(app)

@app.route("/", methods=['GET'])
def index():
    msg = Message("Hi!", sender = 'jambo5167@gmail.com', recipients = ['jirayuth.p@comseven.com'])
    msg.body = "testign test"
    mail.send(msg)
    return "Messege send!"

if __name__ in "__main__":
    app.run()