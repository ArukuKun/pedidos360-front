import React, { useState } from 'react';
import { useMsal, AuthenticatedTemplate, UnauthenticatedTemplate } from '@azure/msal-react';
import { loginRequest } from './AuthConfig';
import './App.css';

const SERVICES = {
  CATALOGO: 'http://<IP_EC2_CATALOGO>:8081/api/catalogo',
  INVENTARIO: 'http://<IP_EC2_INVENTARIO>:8085/api/inventario',
  PEDIDOS: 'http://<IP_EC2_BACKEND>:8080/api/pedidos',
  USUARIOS: 'http://<IP_EC2_USUARIOS>:8082/api/usuarios/perfil',
  PAGOS: 'http://<IP_EC2_PAGOS>:8083/api/pagos/procesar',
  NOTIFICACIONES: 'http://<IP_EC2_NOTIFICACIONES>:8084/api/notificaciones/enviar'
};

function App() {
  const { instance, accounts } = useMsal();
  const [log, setLog] = useState('');
  const [catalogo, setCatalogo] = useState([]);
  const [inventario, setInventario] = useState({});

  const getAccessToken = async () => {
    try {
      const response = await instance.acquireTokenSilent({
        ...loginRequest,
        account: accounts[0]
      });
      return response.accessToken;
    } catch (e) {
      // Modificado para usar Redirect en caso de que falle el token silencioso
      await instance.acquireTokenRedirect(loginRequest);
    }
  };

  const callApi = async (url, method = 'GET', body = null) => {
    try {
      const token = await getAccessToken();
      const headers = {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      };
      
      const config = { method, headers };
      if (body) config.body = JSON.stringify(body);

      const res = await fetch(url, config);
      const data = await res.json();
      setLog(JSON.stringify(data, null, 2));
      return data;
    } catch (err) {
      setLog(`Error al conectar con ${url}: ${err.message}`);
    }
  };

  // Modificados para forzar la redirección completa
  const handleLogin = () => instance.loginRedirect(loginRequest);
  const handleLogout = () => instance.logoutRedirect();

  return (
    <div style={{ padding: '20px', fontFamily: 'sans-serif', maxWidth: '900px', margin: '0 auto' }}>
      <h1>📦 Plataforma Pedidos360</h1>

      <UnauthenticatedTemplate>
        <div style={{ textAlign: 'center', margin: '50px 0' }}>
          <p>Debes iniciar sesión con tu cuenta de Microsoft Entra ID para acceder al sistema.</p>
          <button onClick={handleLogin} style={{ padding: '10px 20px', fontSize: '16px', cursor: 'pointer' }}>
            Iniciar Sesión con Microsoft
          </button>
        </div>
      </UnauthenticatedTemplate>

      <AuthenticatedTemplate>
        <header style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid #ccc', paddingBottom: '10px' }}>
          <span>Bienvenido, <strong>{accounts[0]?.username}</strong></span>
          <button onClick={handleLogout} style={{ padding: '5px 10px', cursor: 'pointer' }}>Cerrar Sesión</button>
        </header>

        <main style={{ marginTop: '20px' }}>
          <h2>Prueba de Microservicios</h2>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '10px' }}>
            <button onClick={async () => setCatalogo(await callApi(SERVICES.CATALOGO))}>
              1. Ver Catálogo (8081)
            </button>
            <button onClick={async () => setInventario(await callApi(SERVICES.INVENTARIO))}>
              2. Ver Inventario (8085)
            </button>
            <button onClick={() => callApi(SERVICES.USUARIOS)}>
              3. Ver Perfil Usuario (8082)
            </button>
            <button onClick={() => callApi(SERVICES.PEDIDOS)}>
              4. Consultar Pedidos (8080)
            </button>
            <button onClick={() => callApi(SERVICES.PAGOS, 'POST', { monto: 15000, numeroTarjeta: '1234567890123456' })}>
              5. Probar Pago OK (8083)
            </button>
            <button onClick={() => callApi(SERVICES.NOTIFICACIONES, 'POST', { destinatario: accounts[0]?.username, mensaje: 'Tu pedido ha sido creado' })}>
              6. Enviar Notificación (8084)
            </button>
          </div>

          <h3 style={{ marginTop: '20px' }}>Respuesta de la API / Log:</h3>
          <pre style={{ background: '#f4f4f4', padding: '15px', borderRadius: '5px', overflowX: 'auto', maxHeight: '300px' }}>
            {log || 'Haz clic en algún botón para probar la integración con los microservicios...'}
          </pre>
        </main>
      </AuthenticatedTemplate>
    </div>
  );
}

export default App;