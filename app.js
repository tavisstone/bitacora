// Variable global para almacenar los datos del Excel
let datosViajes = [];

// Función para simular el inicio de sesión
function iniciarSesion() {
    const user = document.getElementById('username').value;
    const pass = document.getElementById('password').value;

    // Aquí puedes poner una validación simple. 
    // Nota: Como GitHub Pages es estático, la seguridad real se manejaría en un backend,
    // pero para esta etapa, esto funciona para acceder a la interfaz.
    if (user !== "" && pass !== "") {
        document.getElementById('login-section').style.display = 'none';
        document.getElementById('dashboard-section').style.display = 'block';
    } else {
        alert("Por favor, ingrese credenciales válidas.");
    }
}

function cerrarSesion() {
    location.reload(); // Recarga la página para volver al login
}

// Función para cargar y leer el archivo Excel
function cargarExcel(event) {
    const file = event.target.files[0];
    const reader = new FileReader();

    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        
        // Asumimos que los datos están en la primera hoja
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Convertimos el Excel a JSON
        datosViajes = XLSX.utils.sheet_to_json(worksheet);
        
        renderizarTabla(datosViajes);
    };
    reader.readAsArrayBuffer(file);
}

// Función para pintar los datos en el HTML
function renderizarTabla(datos) {
    const tbody = document.querySelector('#viajes-table tbody');
    tbody.innerHTML = '';

    datos.forEach((fila, index) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${fila['ID Viaje'] || 'N/A'}</td>
            <td>${fila['Cliente'] || 'N/A'}</td>
            <td>${fila['Origen'] || 'N/A'}</td>
            <td>${fila['Destino'] || 'N/A'}</td>
            <td>${fila['Estatus Actual'] || 'En espera'}</td>
            <td><button onclick="editarRegistro(${index})">Actualizar</button></td>
        `;
        tbody.appendChild(tr);
    });
}

// Función para generar el archivo de respaldo (Descargar Excel)
function exportarExcel() {
    const ws = XLSX.utils.json_to_sheet(datosViajes);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Bitacora_Respaldo");
    XLSX.writeFile(wb, "Respaldo_Bitacora_GS3.xlsx");
}

function editarRegistro(index) {
    alert("Función de actualización para el registro: " + index);
    // Aquí vincularemos los campos específicos en la siguiente fase
}
