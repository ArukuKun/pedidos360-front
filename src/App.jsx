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
  const [carrito, setCarrito] = useState([]);

  const getAccessToken = async (scopesToRequest) => {
    try {
      const response = await instance.acquireTokenSilent({
        scopes: scopesToRequest || loginRequest.scopes,
        account: accounts[0]
      });
      return response.accessToken;
    } catch (e) {
      try {
        const popupResponse = await instance.acquireTokenPopup({
          scopes: scopesToRequest || loginRequest.scopes
        });
        return popupResponse.accessToken;
      } catch (err) {
        console.error("Error al obtener token:", err);
        return null;
      }
    }
  };

  const callApi = async (url, method = 'GET', body = null, scopes = null) => {
    try {
      const token = await getAccessToken(scopes);
      const headers = {
        'Content-Type': 'application/json'
      };
      if (token) {
        headers['Authorization'] = `Bearer ${token}`;
      }
      
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

  const handleVerCatalogo = async () => {
    const rawCat = await callApi(SERVICES.CATALOGO, 'GET', null, SCOPES.CATALOGO);
    let catList = [];
    try {
      const parsedCat = JSON.parse(rawCat);
      catList = Array.isArray(parsedCat) ? parsedCat : ["Hamburguesa Triple", "Papas Supremas", "Bebida Grande", "Pizza Familiar", "Pastel de Chocolate"];
    } catch {
      catList = ["Hamburguesa Triple", "Papas Supremas", "Bebida Grande", "Pizza Familiar", "Pastel de Chocolate"];
    }
    setCatalogo(catList);

    const rawInv = await callApi(SERVICES.INVENTARIO, 'GET', null, SCOPES.INVENTARIO);
    try {
      const parsedInv = JSON.parse(rawInv);
      if (typeof parsedInv === 'object' && !Array.isArray(parsedInv)) {
        setInventario(parsedInv);
      }
    } catch {
      setInventario({
        "Hamburguesa Triple": 50,
        "Pizza Familiar": 30,
        "Papas Supremas": 100,
        "Pastel de Chocolate": 20,
        "Bebida Grande": 80
      });
    }
  };

  const handleVerInventario = async () => {
    const rawInv = await callApi(SERVICES.INVENTARIO, 'GET', null, SCOPES.INVENTARIO);
    try {
      const parsed = JSON.parse(rawInv);
      if (typeof parsed === 'object') setInventario(parsed);
    } catch {}
  };

  const agregarAlCarrito = (nombreProducto) => {
    const stockActual = inventario[nombreProducto] !== undefined ? inventario[nombreProducto] : 50;
    if (stockActual <= 0) return;

    setInventario(prev => ({
      ...prev,
      [nombreProducto]: (prev[nombreProducto] !== undefined ? prev[nombreProducto] : 50) - 1
    }));

    setCarrito(prev => {
      const existe = prev.find(item => item.nombre === nombreProducto);
      if (existe) {
        return prev.map(item =>
          item.nombre === nombreProducto
            ? { ...item, cantidad: item.cantidad + 1 }
            : item
        );
      }
      return [...prev, { nombre: nombreProducto, cantidad: 1 }];
    });
  };

  const quitarDelCarrito = (nombreProducto) => {
    const itemEnCarrito = carrito.find(item => item.nombre === nombreProducto);
    if (!itemEnCarrito) return;

    setInventario(prev => ({
      ...prev,
      [nombreProducto]: (prev[nombreProducto] || 0) + 1
    }));

    setCarrito(prev => {
      if (itemEnCarrito.cantidad > 1) {
        return prev.map(item =>
          item.nombre === nombreProducto
            ? { ...item, cantidad: item.cantidad - 1 }
            : item
        );
      }
      return prev.filter(item => item.nombre !== nombreProducto);
    });
  };

  const vaciarCarrito = () => {
    setInventario(prev => {
      const nuevoInv = { ...prev };
      carrito.forEach(item => {
        nuevoInv[item.nombre] = (nuevoInv[item.nombre] || 0) + item.cantidad;
      });
      return nuevoInv;
    });
    setCarrito([]);
  };

  const handleCrearPedido = async () => {
    if (carrito.length === 0) return;
    const detallePedido = carrito.map(i => `${i.nombre} (x${i.cantidad})`).join(', ');
    await callApi(SERVICES.PEDIDOS_LOCAL, 'POST', { pedido: detallePedido, items: carrito });
    setCarrito([]);
  };

  const handleLogin = () => instance.loginRedirect(loginRequest);
  const handleLogout = () => instance.logoutRedirect();

  const totalProductosCarrito = carrito.reduce((sum, i) => sum + i.cantidad, 0);

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
            <button onClick={handleVerCatalogo}>1. Ver Catálogo</button>
            <button onClick={handleVerInventario}>2. Ver Inventario</button>
            <button onClick={() => callApi(SERVICES.USUARIOS)}>3. Ver Perfil Usuario</button>
            <button onClick={() => callApi(SERVICES.PEDIDOS_LOCAL)}>4. Consultar Pedidos</button>
            
            <button 
              onClick={handleCrearPedido}
              disabled={carrito.length === 0}
              style={{ 
                backgroundColor: carrito.length > 0 ? '#28a745' : '#6c757d', 
                color: 'white', 
                fontWeight: 'bold',
                cursor: carrito.length > 0 ? 'pointer' : 'not-allowed'
              }}
            >
              🛒 Crear Pedido (RabbitMQ)
            </button>

            <button onClick={() => callApi(SERVICES.PAGOS, 'POST', { monto: 15000, numeroTarjeta: '1234567890123456' })}>5. Probar Pago OK</button>
            <button onClick={() => callApi(SERVICES.NOTIFICACIONES, 'POST', { destinatario: accounts[0]?.username, mensaje: 'Tu pedido ha sido creado' })}>6. Enviar Notificación</button>
          </div>

          <div style={{ marginTop: '20px', padding: '15px', border: '2px solid #28a745', borderRadius: '8px', backgroundColor: '#f4fff6', color: '#000' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '10px' }}>
              <h3 style={{ margin: 0, color: '#155724' }}>
                🛒 Carrito de Compras ({totalProductosCarrito} {totalProductosCarrito === 1 ? 'unidad' : 'unidades'})
              </h3>
              {carrito.length > 0 && (
                <button 
                  onClick={vaciarCarrito}
                  style={{ backgroundColor: '#dc3545', color: 'white', border: 'none', padding: '5px 10px', borderRadius: '4px', cursor: 'pointer' }}
                >
                  Vaciar Carrito
                </button>
              )}
            </div>

            {carrito.length > 0 ? (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                {carrito.map((item) => (
                  <div key={item.nombre} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', backgroundColor: '#fff', padding: '8px 12px', borderRadius: '4px', border: '1px solid #ced4da', color: '#000' }}>
                    <span><strong>{item.nombre}</strong> — Cantidad: <strong>{item.cantidad}</strong></span>
                    <div>
                      <button 
                        onClick={() => agregarAlCarrito(item.nombre)} 
                        disabled={(inventario[item.nombre] !== undefined ? inventario[item.nombre] : 50) <= 0}
                        style={{ backgroundColor: '#28a745', color: 'white', border: 'none', padding: '3px 8px', borderRadius: '3px', cursor: 'pointer', marginRight: '5px' }}
                      >
                        +
                      </button>
                      <button 
                        onClick={() => quitarDelCarrito(item.nombre)} 
                        style={{ backgroundColor: '#ffc107', color: '#000', border: 'none', padding: '3px 8px', borderRadius: '3px', cursor: 'pointer' }}
                      >
                        -
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            ) : (
              <p style={{ margin: 0, color: '#6c757d' }}>El carrito está vacío. Haz clic en "1. Ver Catálogo" y añade productos.</p>
            )}
          </div>

          {catalogo.length > 0 && (
            <div style={{ marginTop: '20px' }}>
              <h3>📦 Catálogo de Productos y Stock:</h3>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: '10px' }}>
                {catalogo.map((prod) => {
                  const stockDisponible = inventario[prod] !== undefined ? inventario[prod] : 50;
                  const tieneStock = stockDisponible > 0;
                  return (
                    <div key={prod} style={{ border: '1px solid #ddd', padding: '12px', borderRadius: '6px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', backgroundColor: '#fff', color: '#000' }}>
                      <div>
                        <div style={{ fontWeight: 'bold', fontSize: '15px' }}>{prod}</div>
                        <div style={{ fontSize: '13px', color: tieneStock ? '#28a745' : '#dc3545', fontWeight: 'bold', marginTop: '4px' }}>
                          {tieneStock ? `Stock disponible: ${stockDisponible}` : '❌ Agotado'}
                        </div>
                      </div>
                      <button 
                        onClick={() => agregarAlCarrito(prod)}
                        disabled={!tieneStock}
                        style={{ 
                          backgroundColor: tieneStock ? '#007bff' : '#6c757d', 
                          color: 'white', 
                          border: 'none', 
                          padding: '8px 12px', 
                          cursor: tieneStock ? 'pointer' : 'not-allowed', 
                          borderRadius: '4px',
                          fontWeight: 'bold'
                        }}
                      >
                        {tieneStock ? '➕ Añadir al Carrito' : 'Agotado'}
                      </button>
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          <h3 style={{ marginTop: '20px' }}>Respuesta de la API / Log:</h3>
          <pre style={{ background: '#f4f4f4', color: '#000', padding: '15px', borderRadius: '5px', overflowX: 'auto', maxHeight: '300px' }}>
            {log || 'Haz clic en algún botón para probar la integración con los microservicios...'}
          </pre>
        </main>
      </AuthenticatedTemplate>
    </div>
  );
}

export default App;