from flask import Flask, request, jsonify, render_template
import pandas as pd
import subprocess
import os

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
        if not os.path.exists(ruta_excel):
            raise FileNotFoundError("El archivo Excel no existe en la ruta especificada.")
        
        df = pd.read_excel(ruta_excel)
        print(df.head())  # Imprimir el contenido del DataFrame para verificarlo

        expected_columns = ['usuario', 'contraseña']
        if not all(column in df.columns for column in expected_columns):
            raise ValueError("El archivo Excel no contiene las columnas esperadas.")

        if df.empty:
            raise ValueError("El archivo Excel está vacío.")

        data = df.to_dict(orient='records')
        return jsonify(data)

    except FileNotFoundError as fnf_error:
        return jsonify({"error": str(fnf_error)}), 404
    except ValueError as val_error:
        return jsonify({"error": str(val_error)}), 400
    except pd.errors.ExcelFileError as excel_error:
        return jsonify({"error": "Error al leer el archivo Excel: " + str(excel_error)}), 500
    except Exception as e:
        return jsonify({"error": "Error inesperado: " + str(e)}), 500

@app.route('/guardar_excel', methods=['POST'])
def guardar_excel():
    try:
        data = request.json
        df = pd.DataFrame(data)
        with pd.ExcelWriter(ruta_excel) as writer:
            df.to_excel(writer, index=False)
        return jsonify({"message": "Datos guardados exitosamente."})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/crear_usuarios_excel', methods=['POST'])
def crear_usuarios_excel():
    script_path_create = 'C:\\Users\\Santiago\\Documents\\crear_usuarios_excel.ps1'
    try:
        result = subprocess.run(["powershell", "-File", script_path_create], capture_output=True, text=True)
        output = result.stdout
        error = result.stderr
        if result.returncode != 0:
            return jsonify({"output": output, "error": error}), 500
        return jsonify({"output": output, "error": error})
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
    try:
        output = subprocess.check_output(
            ["powershell", "C:\\Users\\Santiago\\Documents\\eliminar_permisos.ps1", nombre_usuario, nombre_carpeta],
            universal_newlines=True
        )
    except subprocess.CalledProcessError as e:
        output = f"Error al ejecutar el script: {e.output}"
    
    return render_template('output.html', output=output)

if __name__ == '__main__':
    app.run(debug=True)
