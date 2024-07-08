from flask import Flask, request, jsonify, render_template
import pandas as pd
import subprocess

app = Flask(__name__)

# Ruta al archivo Excel
ruta_excel = "C:\\Users\\Santiago\\Documents\\usuarios.xlsx"

# Ruta a los scripts de PowerShell
script_path_create = "C:\\Users\\Santiago\\Documents\\crear_usuarios_excel.ps1"
script_path_delete = "C:\\Users\\Santiago\\Documents\\eliminar_usuarios_excel.ps1"
script_path_create_user = "C:\\Users\\Santiago\\Documents\\crear_usuario.ps1"
script_path_delete_user = "C:\\Users\\Santiago\\Documents\\eliminar_usuario.ps1"
script_path_create_folder = "C:\\Users\\Santiago\\Documents\\crear_carpeta.ps1"
script_path_enable_permissions = "C:\\Users\\Santiago\\Documents\\habilitar_permisos.ps1"
script_path_disable_permissions = "C:\\Users\\Santiago\\Documents\\eliminar_permisos.ps1"

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/crear_usuario', methods=['POST'])
def crear_usuario():
    try:
        nombre = request.json['nombre']
        contrasena = request.json['contrasena']
        result = subprocess.run(["powershell", "-File", script_path_create_user, "-Nombre", nombre, "-Contrasena", contrasena], capture_output=True, text=True)
        return jsonify({"output": result.stdout, "error": result.stderr})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/eliminar_usuario', methods=['POST'])
def eliminar_usuario():
    try:
        nombre = request.json['nombre']
        result = subprocess.run(["powershell", "-File", script_path_delete_user, "-Nombre", nombre], capture_output=True, text=True)
        return jsonify({"output": result.stdout, "error": result.stderr})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/leer_excel', methods=['GET'])
def leer_excel():
    try:
        df = pd.read_excel(ruta_excel)
        data = df.to_dict(orient='records')
        return jsonify(data)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/guardar_excel', methods=['POST'])
def guardar_excel():
    try:
        data = request.json
        df = pd.DataFrame(data)
        df.to_excel(ruta_excel, index=False)
        return jsonify({"message": "Datos guardados exitosamente."})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/crear_usuarios_excel', methods=['POST'])
def crear_usuarios_excel():
    try:
        result = subprocess.run(["powershell", "-File", script_path_create], capture_output=True, text=True)
        return jsonify({"output": result.stdout, "error": result.stderr})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/eliminar_usuarios_excel', methods=['POST'])
def eliminar_usuarios_excel():
    try:
        result = subprocess.run(["powershell", "-File", script_path_delete], capture_output=True, text=True)
        return jsonify({"output": result.stdout, "error": result.stderr})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/crear_carpeta', methods=['POST'])
def crear_carpeta():
    try:
        nombre_usuario = request.json['nombre_usuario']
        nombre_carpeta = request.json['nombre_carpeta']
        result = subprocess.run(["powershell", "-File", script_path_create_folder, "-NombreUsuario", nombre_usuario, "-NombreCarpeta", nombre_carpeta], capture_output=True, text=True)
        return jsonify({"output": result.stdout, "error": result.stderr})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/habilitar_permisos', methods=['POST'])
def habilitar_permisos():
    nombre_usuario = request.form['nombre_usuario']
    nombre_carpeta = request.form['nombre_carpeta']
    output = subprocess.check_output(
        ["powershell", "C:\\Users\\Santiago\\Documents\\habilitar_permisos.ps1", nombre_usuario, nombre_carpeta],
        universal_newlines=True
    )
    return render_template('output.html', output=output)

@app.route('/eliminar_permisos', methods=['POST'])
def eliminar_permisos():
    nombre_usuario = request.form['nombre_usuario']
    nombre_carpeta = request.form['nombre_carpeta']
    output = subprocess.check_output(
        ["powershell", "C:\\Users\\Santiago\\Documents\\eliminar_permisos.ps1", nombre_usuario, nombre_carpeta],
        universal_newlines=True
    )
    return render_template('output.html', output=output)

if __name__ == '__main__':
    app.run(debug=True)
