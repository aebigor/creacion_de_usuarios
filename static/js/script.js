document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('leer-excel').addEventListener('click', () => {
        fetch('/leer_excel')
            .then(response => {
                if (!response.ok) {
                    throw new Error('Ocurrió un problema al leer el archivo Excel.');
                }
                return response.json();
            })
            .then(data => {
                const tbody = document.getElementById('tabla-usuarios').getElementsByTagName('tbody')[0];
                tbody.innerHTML = '';
                if (data.length === 0) {
                    document.getElementById('output').textContent = 'No hay datos en el archivo Excel.';
                    return;
                }
                data.forEach(fila => {
                    const tr = document.createElement('tr');
                    const tdUsuario = document.createElement('td');
                    const inputUsuario = document.createElement('input');
                    inputUsuario.type = 'text';
                    inputUsuario.value = fila.usuario || 'N/A';
                    tdUsuario.appendChild(inputUsuario);
                    const tdContrasena = document.createElement('td');
                    const inputContrasena = document.createElement('input');
                    inputContrasena.type = 'text';
                    inputContrasena.value = fila.contraseña || 'N/A';
                    tdContrasena.appendChild(inputContrasena);
                    tr.appendChild(tdUsuario);
                    tr.appendChild(tdContrasena);
                    tbody.appendChild(tr);
                });
                document.getElementById('output').textContent = 'Datos del Excel cargados correctamente.';
            })
            .catch(error => {
                document.getElementById('output').textContent = 'Error: ' + error.message;
                console.error('Error:', error);
            });
    });
    document.getElementById('guardar-excel').addEventListener('click', () => {
        const tbody = document.getElementById('tabla-usuarios').getElementsByTagName('tbody')[0];
        const filas = Array.from(tbody.getElementsByTagName('tr'));
        const data = filas.map(fila => {
            const celdas = fila.getElementsByTagName('td');
            return {
                usuario: celdas[0].getElementsByTagName('input')[0].value,
                contraseña: celdas[1].getElementsByTagName('input')[0].value
            };
        });

        fetch('/guardar_excel', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(data)
        })
            .then(response => response.json())
            .then(data => {
                document.getElementById('output').textContent = data.message || data.error;
            })
            .catch(error => {
                document.getElementById('output').textContent = 'Error: ' + error.message;
                console.error('Error:', error);
            });
    });
});

document.getElementById('crear-usuarios-excel').addEventListener('click', () => {
    fetch('/crear_usuarios_excel', { method: 'POST' })
        .then(response => response.json())
        .then(data => {
            document.getElementById('output').textContent = `Output: ${data.output}\nError: ${data.error}`;
        })
        .catch(error => console.error('Error:', error));
});

document.getElementById('eliminar-usuarios-excel').addEventListener('click', () => {
    fetch('/eliminar_usuarios_excel', { method: 'POST' })
        .then(response => response.json())
        .then(data => {
            document.getElementById('output').textContent = `Output: ${data.output}\nError: ${data.error}`;
        })
        .catch(error => console.error('Error:', error));
});

document.getElementById('form-crear-usuario').addEventListener('submit', (e) => {
    e.preventDefault();
    const nombre = document.getElementById('nombre').value;
    const contrasena = document.getElementById('contrasena').value;

    fetch('/crear_usuario', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ nombre, contrasena })
    })
        .then(response => response.json())
        .then(data => {
            document.getElementById('output').textContent = `Output: ${data.output}\nError: ${data.error}`;
        })
        .catch(error => console.error('Error:', error));
});
document.getElementById('form-eliminar-usuario').addEventListener('submit', (e) => {
    e.preventDefault();
    const nombre = document.getElementById('nombre-eliminar').value;

    fetch('/eliminar_usuario', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ nombre })
    })
        .then(response => response.json())
        .then(data => {
            document.getElementById('output').textContent = `Output: ${data.output}\nError: ${data.error}`;
        })
        .catch(error => console.error('Error:', error));
});
document.getElementById('form-crear-carpeta').addEventListener('submit', (e) => {
    e.preventDefault();
    const nombre_usuario = document.getElementById('nombre-usuario-carpeta').value;
    const nombre_carpeta = document.getElementById('nombre-carpeta').value;

    fetch('/crear_carpeta', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ nombre_usuario, nombre_carpeta })
    })
        .then(response => response.json())
        .then(data => {
            document.getElementById('output').textContent = `Output: ${data.output}\nError: ${data.error}`;
        })
        .catch(error => console.error('Error:', error));
});

document.getElementById('form-habilitar-permisos').addEventListener('submit', (e) => {
    e.preventDefault();
    const nombre_usuario = document.getElementById('nombre-usuario-permisos').value;
    const nombre_carpeta = document.getElementById('nombre-carpeta-permisos').value;

    fetch('/habilitar_permisos', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ nombre_usuario, nombre_carpeta })
    })
        .then(response => response.json())
        .then(data => {
            document.getElementById('output').textContent = `Output: ${data.output}\nError: ${data.error}`;
        })
        .catch(error => console.error('Error:', error));
});

document.getElementById('form-eliminar-permisos').addEventListener('submit', (e) => {
    e.preventDefault();
    const nombre_usuario = document.getElementById('nombre-usuario-eliminar-permisos').value;
    const nombre_carpeta = document.getElementById('nombre-carpeta-eliminar-permisos').value;

    fetch('/eliminar_permisos', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ nombre_usuario, nombre_carpeta })
    })
        .then(response => response.json())
        .then(data => {
            document.getElementById('output').textContent = `Output: ${data.output}\nError: ${data.error}`;
        })
        .catch(error => console.error('Error:', error));
});
document.addEventListener('DOMContentLoaded', () => {
document.getElementById('leer-excel').addEventListener('click', () => {
    fetch('/leer_excel')
        .then(response => {
            if (!response.ok) {
                throw new Error('Ocurrió un problema al leer el archivo Excel.');
            }
            return response.json();
        })
        .then(data => {
            const tbody = document.getElementById('tabla-usuarios').getElementsByTagName('tbody')[0];
            tbody.innerHTML = '';
            data.forEach(fila => {
                const tr = document.createElement('tr');
                const tdUsuario = document.createElement('td');
                tdUsuario.textContent = fila.Usuario;
                const tdContrasena = document.createElement('td');
                tdContrasena.textContent = fila.Contrasena;
                tr.appendChild(tdUsuario);
                tr.appendChild(tdContrasena);
                tbody.appendChild(tr);
            });
            document.getElementById('output').textContent = 'Datos del Excel cargados correctamente.';
        })
        .catch(error => {
            document.getElementById('output').textContent = 'Error: ' + error.message;
            console.error('Error:', error);
        });
});
});