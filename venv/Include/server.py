from flask import Flask, request, jsonify
import os, sys, socket, parsingxls
oss = os.name
app = Flask(__name__)

#Роут предоставляет доступ к методу run_operating с параметром = '-s' из класса Run в файле ves.py
@app.route("/get_svin", methods=['POST', 'GET'])
def svin():
    if request.method == 'GET':
        if parsingxls.verify_ip(request.environ['REMOTE_ADDR']) == 1:
            parsingxls.run('-s')
            res = open('tmp/svin_res.r', mode='r+', encoding='utf-8')
            return res.read(), {'Content-Type': 'text/json'}
        else:
            return "Access denied"
#Роут предоставляет доступ к методу run_operating с параметром = '-k' из класса Run в файле ves.py
@app.route("/get_korm", methods=['POST', 'GET'])
def korm():
    if request.method == 'GET':
        if parsingxls.verify_ip(request.environ['REMOTE_ADDR']) == 1:
            parsingxls.run('-k')
            res = open('tmp/korm_res.r', mode='r+', encoding='utf-8')
            return res.read(),{'Content-Type': 'text/json'}
        else:
            return "Access denied"
@app.route("/cleantemp", methods=['POST', 'GET'])
def cleantemp():
    if request.method == 'GET':
        if access.verify_ip(request.environ['REMOTE_ADDR']) == 1:
            parsingxls.cleaner()

@app.route("/view_my_ip", methods=['POST', 'GET'])
def viewip():
    if request.method == 'GET':
        return request.environ['REMOTE_ADDR']
#В зависимости от операционной системе стартует веб сервер при выполнении скрипта server.py
if oss == 'posix':
    if __name__ == "__main__":
        app.run('192.168.100.34', port=8000)
if oss == 'nt':
    if __name__ == "__main__":
        app.run(host=socket.gethostbyname(socket.getfqdn()),port=8000)