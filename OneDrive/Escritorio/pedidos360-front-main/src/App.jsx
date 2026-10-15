import React, { useState } from 'react';
import { useMsal, AuthenticatedTemplate, UnauthenticatedTemplate } from '@azure/msal-react';
import { loginRequest } from './auth/AuthConfig';
import './App.css';

const SERVICES = {
  CATALOGO: 'https://j76lk49uq2.execute-api.us-east-1.amazonaws.com/v1/catalogo',
  INVENTARIO: 'https://j76lk49uq2.execute-api.us-east-1.amazonaws.com/v1/inventario',
  PEDIDOS_LOCAL: 'http://localhost:8080/api/pedidos',
  PEDIDOS: 'https://j76lk49uq2.execute-api.us-east-1.amazonaws.com/v1/pedidos',
  USUARIOS: 'https://j76lk49uq2.execute-api.us-east-1.amazonaws.com/v1/usuarios/perfil',
  PAGOS: 'https://j76lk49uq2.execute-api.us-east-1.amazonaws.com/v1/pagos/procesar',
  NOTIFICACIONES: 'https://j76lk49uq2.execute-api.us-east-1.amazonaws.com/v1/notificaciones'
};

const SCOPES = {
  CATALOGO: ["api://5d7ead55-12ca-4dcd-97e8-5b5e81f77f4b/Catalogo.Read"],
  INVENTARIO: ["api://1f835a60-98e3-4fde-b87f-0c64cd1e7e46/Inventario.Read"]
};

function App() {
  const { instance, accounts } = useMsal();
  const [log, setLog] = useState('');
  const [catalogo, setCatalogo] = useState([]);
  const [inventario, setInventario] = useState({});

  const getAccessToken = async (scopesToRequest) => {
    try {
      const response = await instance.acquireTokenSilent({
        scopes: scopesToRequest || loginRequest.scopes,
        account: accounts[0]
      });
      return response.accessToken;
    } catch (e) {
      await instance.acquireTokenRedirect({
        scopes: scopesToRequest || loginRequest.scopes
      });
    }
  };

  const callApi = async (url, method = 'GET', body = null, scopes = null) => {
    try {
      const token = await getAccessToken(scopes);
      const headers = {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      };
      
      const config = { method, headers };
      if (body) config.body = JSON.stringify(body);

      const res = await fetch(url, config);
      const data = await res.text(); 
      
      try {
        setLog(JSON.stringify(JSON.parse(data), null, 2));
      } catch {
        setLog(data);
      }
      return data;
    } catch (err) {
      setLog(`Error al conectar con ${url}: ${err.message}`);
    }
  };

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
            <button onClick={async () => setCatalogo(await callApi(SERVICES.CATALOGO, 'GET', null, SCOPES.CATALOGO))}>1. Ver Catálogo</button>
            <button onClick={async () => setInventario(await callApi(SERVICES.INVENTARIO, 'GET', null, SCOPES.INVENTARIO))}>2. Ver Inventario</button>
            <button onClick={() => callApi(SERVICES.USUARIOS)}>3. Ver Perfil Usuario</button>
            <button onClick={() => callApi(SERVICES.PEDIDOS_LOCAL)}>4. Consultar Pedidos</button>
            
            {/* NUEVO BOTÓN: Crea el pedido mediante POST disparando la cola de RabbitMQ */}
            <button 
              onClick={() => callApi(SERVICES.PEDIDOS_LOCAL, 'POST', 'Teclado Mecánico RGB')}
              style={{ backgroundColor: '#28a745', color: 'white', fontWeight: 'bold' }}
            >
              🛒 Crear Pedido (RabbitMQ)
            </button>

            <button onClick={() => callApi(SERVICES.PAGOS, 'POST', { monto: 15000, numeroTarjeta: '1234567890123456' })}>5. Probar Pago OK</button>
            <button onClick={() => callApi(SERVICES.NOTIFICACIONES, 'POST', { destinatario: accounts[0]?.username, mensaje: 'Tu pedido ha sido creado' })}>6. Enviar Notificación</button>
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