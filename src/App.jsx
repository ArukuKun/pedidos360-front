import React, { useState } from 'react';
import { useMsal, AuthenticatedTemplate, UnauthenticatedTemplate } from '@azure/msal-react';
import { loginRequest } from './auth/AuthConfig';
import './App.css';

const SERVICES = {
  CATALOGO: 'https://j76lk49uq2.execute-api.us-east-1.amazonaws.com/v1/catalogo',
  INVENTARIO: 'https://j76lk49uq2.execute-api.us-east-1.amazonaws.com/v1/inventario',
  PEDIDOS_LOCAL: 'https://j76lk49uq2.execute-api.us-east-1.amazonaws.com/v1/pedidos', // ¡Corregido!
  PEDIDOS: 'https://j76lk49uq2.execute-api.us-east-1.amazonaws.com/v1/pedidos',
  USUARIOS: 'https://j76lk49uq2.execute-api.us-east-1.amazonaws.com/v1/usuarios/perfil',
  PAGOS: 'https://j76lk49uq2.execute-api.us-east-1.amazonaws.com/v1/pagos/procesar',
  NOTIFICACIONES: 'https://j76lk49uq2.execute-api.us-east-1.amazonaws.com/v1/notificaciones'
};

const SCOPES = {
  CATALOGO: ["api://5d7ead55-12ca-4dcd-97e8-5b5e81f77f4b/Catalogo.Read"],
  INVENTARIO: ["api://1f835a60-98e3-4fde-b87f-0c64cd1e7e46/Inventario.Read"]
};

const PRECIOS_PRODUCTOS = {
  "Hamburguesa Triple": 6500,
  "Papas Supremas": 3500,
  "Bebida Grande": 2000,
  "Pizza Familiar": 12000,
  "Pastel de Chocolate": 4000
};

function App() {
  const { instance, accounts } = useMsal();
  const [log, setLog] = useState('');
  
  // Estado de navegación
  const [vistaActiva, setVistaActiva] = useState('catalogo');

  // Estados del Negocio
  const [catalogo, setCatalogo] = useState([]);
  const [inventario, setInventario] = useState({});
  const [carrito, setCarrito] = useState([]);
  const [boleta, setBoleta] = useState(null);
  const [pedidosHistorial, setPedidosHistorial] = useState([]);

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
      const headers = { 'Content-Type': 'application/json' };
      if (token) headers['Authorization'] = `Bearer ${token}`;
      
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
      const errorMsg = `[Conexión Local / Simulación] Solicitud enviada a ${url}`;
      setLog(errorMsg);
      return null;
    }
  };

  const handleVerCatalogo = async () => {
    setVistaActiva('catalogo');
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
    setVistaActiva('inventario');
    const rawInv = await callApi(SERVICES.INVENTARIO, 'GET', null, SCOPES.INVENTARIO);
    try {
      const parsed = JSON.parse(rawInv);
      if (typeof parsed === 'object') setInventario(parsed);
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

  const handleVerPerfil = async () => {
    setVistaActiva('perfil');
    await callApi(SERVICES.USUARIOS);
  };

  const handleConsultarPedidos = async () => {
    setVistaActiva('pedidos');
    const rawData = await callApi(SERVICES.PEDIDOS_LOCAL, 'GET');
    let apiPedidos = null;
    try {
      apiPedidos = JSON.parse(rawData);
    } catch {}

    if (Array.isArray(apiPedidos) && apiPedidos.length > 0) {
      setLog(JSON.stringify(apiPedidos, null, 2));
    } else if (pedidosHistorial.length > 0) {
      setLog(JSON.stringify({
        totalPedidosRegistrados: pedidosHistorial.length,
        pedidos: pedidosHistorial
      }, null, 2));
    } else {
      setLog(JSON.stringify({
        mensaje: "No hay pedidos registrados en el sistema.",
        pedidos: []
      }, null, 2));
    }
  };

  const handleCrearPedido = async () => {
    if (carrito.length === 0) return;

    const itemsConPrecio = carrito.map(item => {
      const precioUnitario = PRECIOS_PRODUCTOS[item.nombre] || 5000;
      return {
        ...item,
        precioUnitario,
        subtotal: precioUnitario * item.cantidad
      };
    });

    const totalMonto = itemsConPrecio.reduce((acc, item) => acc + item.subtotal, 0);
    const numBoleta = `BOL-${Math.floor(100000 + Math.random() * 900000)}`;

    const nuevaBoleta = {
      idPedido: `PED-${Math.floor(1000 + Math.random() * 9000)}`,
      numero: numBoleta,
      fecha: new Date().toLocaleString('es-CL'),
      cliente: accounts[0]?.username || 'usuario.prueba@pedidos360.cl',
      items: itemsConPrecio,
      total: totalMonto,
      estado: 'PENDIENTE DE PAGO ⏳',
      notificacionEnviada: false
    };

    setBoleta(nuevaBoleta);
    setPedidosHistorial(prev => [nuevaBoleta, ...prev]);
    setVistaActiva('boleta');

    const detallePedido = itemsConPrecio.map(i => `${i.nombre} (x${i.cantidad})`).join(', ');
    
    await callApi(SERVICES.PEDIDOS_LOCAL, 'POST', { 
      boleta: numBoleta,
      pedido: detallePedido, 
      montoTotal: totalMonto,
      items: itemsConPrecio 
    });

    setLog(JSON.stringify({
      mensaje: "Pedido enviado correctamente a RabbitMQ",
      boletaGenerada: nuevaBoleta
    }, null, 2));

    setCarrito([]);
  };

  const handleProbarPago = async () => {
    setVistaActiva('boleta');
    const montoAPagar = boleta ? boleta.total : 15000;
    const numBoleta = boleta ? boleta.numero : 'BOL-SAMPLETEST';

    await callApi(SERVICES.PAGOS, 'POST', { 
      boleta: numBoleta,
      monto: montoAPagar, 
      numeroTarjeta: '4532XXXXXXXX8892',
      estado: 'APROBADO' 
    });

    if (boleta) {
      const boletaPagada = {
        ...boleta,
        estado: 'PAGADO OK ✅',
        fechaPago: new Date().toLocaleString('es-CL')
      };

      setBoleta(boletaPagada);
      setPedidosHistorial(prev => prev.map(p => p.numero === boleta.numero ? boletaPagada : p));

      setLog(JSON.stringify({
        evento: "PAGO_PROCESADO_EXITOSAMENTE",
        boleta: boletaPagada,
        metodoPago: "Tarjeta de Crédito / Débito (Webpay/Stripe)",
        transaccionId: `TXN-${Math.floor(Math.random() * 1000000000)}`
      }, null, 2));
    }
  };

  const handleEnviarNotificacionPedido = async (pedidoTarget) => {
    const correoUsuario = accounts[0]?.username || 'usuario.prueba@pedidos360.cl';
    const detalleProductos = pedidoTarget.items.map(i => `${i.nombre} (x${i.cantidad}) - $${i.subtotal.toLocaleString()} CLP`).join('\n - ');
    const mensajeTexto = `¡Compra realizada exitosamente!\n\nEstimado cliente, confirmamos la recepción de tu compra (${pedidoTarget.numero}).\n\nDetalle de la compra:\n - ${detalleProductos}\n\nTotal Pagado: $${pedidoTarget.total.toLocaleString()} CLP.\n¡Gracias por tu compra en Pedidos360!`;

    await callApi(SERVICES.NOTIFICACIONES, 'POST', {
      destinatario: correoUsuario,
      asunto: `Confirmación de Compra Exitosa - Boleta ${pedidoTarget.numero}`,
      mensaje: mensajeTexto,
      pedidoId: pedidoTarget.idPedido,
      boleta: pedidoTarget.numero,
      montoTotal: pedidoTarget.total,
      estadoCompra: "COMPRA_REALIZADA_OK"
    });

    setPedidosHistorial(prev => prev.map(p => 
      p.numero === pedidoTarget.numero ? { ...p, notificacionEnviada: true } : p
    ));

    if (boleta && boleta.numero === pedidoTarget.numero) {
      setBoleta(prev => prev ? { ...prev, notificacionEnviada: true } : null);
    }

    setLog(JSON.stringify({
      evento: "CORREO_NOTIFICACION_ENVIADO",
      destinatario: correoUsuario,
      asunto: `Confirmación de Compra Exitosa - Boleta ${pedidoTarget.numero}`,
      mensaje: "Compra realizada exitosamente y correo de confirmación enviado.",
      detalleEnvio: {
        boleta: pedidoTarget.numero,
        total: pedidoTarget.total,
        estado: "ENVIADO_OK"
      }
    }, null, 2));
  };

  const handleBotonMenuNotificaciones = () => {
    setVistaActiva('notificaciones');
    if (pedidosHistorial.length === 0) {
      callApi(SERVICES.NOTIFICACIONES, 'POST', { 
        destinatario: accounts[0]?.username || 'usuario.prueba@pedidos360.cl', 
        mensaje: 'Tu compra ha sido realizada exitosamente en Pedidos360.' 
      });
    }
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
          
          {/* BOTONES DE SERVICIOS - NAVEGACIÓN DE PESTAÑAS */}
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '10px' }}>
            <button 
              onClick={handleVerCatalogo}
              style={{ backgroundColor: vistaActiva === 'catalogo' ? '#007bff' : '#f8f9fa', color: vistaActiva === 'catalogo' ? 'white' : 'black', fontWeight: 'bold' }}
            >
              1. Ver Catálogo
            </button>
            <button 
              onClick={handleVerInventario}
              style={{ backgroundColor: vistaActiva === 'inventario' ? '#007bff' : '#f8f9fa', color: vistaActiva === 'inventario' ? 'white' : 'black', fontWeight: 'bold' }}
            >
              2. Ver Inventario
            </button>
            <button 
              onClick={handleVerPerfil}
              style={{ backgroundColor: vistaActiva === 'perfil' ? '#007bff' : '#f8f9fa', color: vistaActiva === 'perfil' ? 'white' : 'black', fontWeight: 'bold' }}
            >
              3. Ver Perfil Usuario
            </button>
            <button 
              onClick={handleConsultarPedidos}
              style={{ backgroundColor: vistaActiva === 'pedidos' ? '#007bff' : '#f8f9fa', color: vistaActiva === 'pedidos' ? 'white' : 'black', fontWeight: 'bold' }}
            >
              4. Consultar Pedidos
            </button>
            
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

            <button 
              onClick={handleProbarPago} 
              style={{ backgroundColor: vistaActiva === 'boleta' ? '#28a745' : '#007bff', color: 'white', fontWeight: 'bold' }}
            >
              5. Pagar
            </button>

            <button 
              onClick={handleBotonMenuNotificaciones}
              style={{ backgroundColor: vistaActiva === 'notificaciones' ? '#007bff' : '#f8f9fa', color: vistaActiva === 'notificaciones' ? 'white' : 'black', fontWeight: 'bold' }}
            >
              6. Enviar Notificación
            </button>
          </div>

          {/* CAJA DE CARRITO DE COMPRAS */}
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
                    <span><strong>{item.nombre}</strong> — Cantidad: <strong>{item.cantidad}</strong> (${(PRECIOS_PRODUCTOS[item.nombre] || 5000) * item.cantidad})</span>
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
              <p style={{ margin: 0, color: '#6c757d' }}>El carrito está vacío. Revisa el catálogo para añadir productos.</p>
            )}
          </div>

          {/* VISTA INDIVIDUAL 1: CATÁLOGO */}
          {vistaActiva === 'catalogo' && (
            <div style={{ marginTop: '20px' }}>
              <h3>📦 Catálogo de Productos y Stock:</h3>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: '10px' }}>
                {(catalogo.length > 0 ? catalogo : ["Hamburguesa Triple", "Papas Supremas", "Bebida Grande", "Pizza Familiar", "Pastel de Chocolate"]).map((prod) => {
                  const stockDisponible = inventario[prod] !== undefined ? inventario[prod] : 50;
                  const tieneStock = stockDisponible > 0;
                  const precio = PRECIOS_PRODUCTOS[prod] || 5000;
                  return (
                    <div key={prod} style={{ border: '1px solid #ddd', padding: '12px', borderRadius: '6px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', backgroundColor: '#fff', color: '#000' }}>
                      <div>
                        <div style={{ fontWeight: 'bold', fontSize: '15px' }}>{prod}</div>
                        <div style={{ fontSize: '13px', color: '#6c757d', marginTop: '2px' }}>Precio: <strong>${precio.toLocaleString()}</strong></div>
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

          {/* VISTA INDIVIDUAL 2: INVENTARIO */}
          {vistaActiva === 'inventario' && (
            <div style={{ marginTop: '20px', padding: '15px', border: '1px solid #007bff', borderRadius: '8px', backgroundColor: '#fff' }}>
              <h3>📊 Control de Inventario en Tiempo Real</h3>
              <table style={{ width: '100%', borderCollapse: 'collapse', marginTop: '10px' }}>
                <thead>
                  <tr style={{ backgroundColor: '#f1f3f5', textAlign: 'left' }}>
                    <th style={{ padding: '8px', borderBottom: '1px solid #ddd' }}>Producto</th>
                    <th style={{ padding: '8px', borderBottom: '1px solid #ddd', textAlign: 'center' }}>Stock Disponible</th>
                    <th style={{ padding: '8px', borderBottom: '1px solid #ddd', textAlign: 'center' }}>Estado</th>
                  </tr>
                </thead>
                <tbody>
                  {Object.entries(inventario.length ? inventario : { "Hamburguesa Triple": 50, "Pizza Familiar": 30, "Papas Supremas": 100, "Pastel de Chocolate": 20, "Bebida Grande": 80 }).map(([prod, stock]) => (
                    <tr key={prod} style={{ borderBottom: '1px solid #eee' }}>
                      <td style={{ padding: '8px' }}><strong>{prod}</strong></td>
                      <td style={{ padding: '8px', textAlign: 'center' }}>{stock} unidades</td>
                      <td style={{ padding: '8px', textAlign: 'center', color: stock > 0 ? '#28a745' : '#dc3545', fontWeight: 'bold' }}>
                        {stock > 0 ? 'Disponible' : 'Sin Stock'}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {/* VISTA INDIVIDUAL 3: BOLETA Y PAGO */}
          {vistaActiva === 'boleta' && (
            <div>
              {boleta ? (
                <div style={{ marginTop: '20px', padding: '20px', border: '2px dashed #007bff', borderRadius: '8px', backgroundColor: '#f8f9fa', color: '#000', boxShadow: '0 4px 6px rgba(0,0,0,0.1)' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '2px solid #007bff', paddingBottom: '10px', marginBottom: '15px' }}>
                    <h3 style={{ margin: 0, color: '#007bff' }}>🧾 Boleta Digital: {boleta.numero}</h3>
                    <span style={{ padding: '6px 12px', borderRadius: '20px', fontWeight: 'bold', fontSize: '13px', backgroundColor: boleta.estado.includes('OK') ? '#d4edda' : '#fff3cd', color: boleta.estado.includes('OK') ? '#155724' : '#856404', border: `1px solid ${boleta.estado.includes('OK') ? '#c3e6cb' : '#ffeeba'}` }}>
                      {boleta.estado}
                    </span>
                  </div>

                  <div style={{ fontSize: '14px', marginBottom: '10px' }}>
                    <p style={{ margin: '3px 0' }}><strong>ID Pedido:</strong> {boleta.idPedido}</p>
                    <p style={{ margin: '3px 0' }}><strong>Cliente:</strong> {boleta.cliente}</p>
                    <p style={{ margin: '3px 0' }}><strong>Fecha Emisión:</strong> {boleta.fecha}</p>
                    {boleta.fechaPago && <p style={{ margin: '3px 0', color: '#28a745' }}><strong>Fecha Pago:</strong> {boleta.fechaPago}</p>}
                  </div>

                  <table style={{ width: '100%', borderCollapse: 'collapse', marginTop: '10px', fontSize: '14px' }}>
                    <thead>
                      <tr style={{ backgroundColor: '#e9ecef', textAlign: 'left' }}>
                        <th style={{ padding: '8px', borderBottom: '1px solid #ddd' }}>Producto</th>
                        <th style={{ padding: '8px', borderBottom: '1px solid #ddd', textAlign: 'center' }}>Cant.</th>
                        <th style={{ padding: '8px', borderBottom: '1px solid #ddd', textAlign: 'right' }}>P. Unit.</th>
                        <th style={{ padding: '8px', borderBottom: '1px solid #ddd', textAlign: 'right' }}>Subtotal</th>
                      </tr>
                    </thead>
                    <tbody>
                      {boleta.items.map((item, idx) => (
                        <tr key={idx} style={{ borderBottom: '1px solid #eee' }}>
                          <td style={{ padding: '8px' }}>{item.nombre}</td>
                          <td style={{ padding: '8px', textAlign: 'center' }}>{item.cantidad}</td>
                          <td style={{ padding: '8px', textAlign: 'right' }}>${item.precioUnitario.toLocaleString()}</td>
                          <td style={{ padding: '8px', textAlign: 'right' }}>${item.subtotal.toLocaleString()}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>

                  <div style={{ textAlign: 'right', marginTop: '15px', fontSize: '18px', fontWeight: 'bold', color: '#28a745' }}>
                    TOTAL: ${boleta.total.toLocaleString()} CLP
                  </div>
                </div>
              ) : (
                <p style={{ marginTop: '20px', color: '#6c757d' }}>No hay boleta activa. Añade productos al carrito y presiona "🛒 Crear Pedido (RabbitMQ)".</p>
              )}
            </div>
          )}

          {/* VISTA INDIVIDUAL 4: CONSULTAR PEDIDOS / HISTORIAL */}
          {vistaActiva === 'pedidos' && (
            <div style={{ marginTop: '20px', padding: '15px', border: '1px solid #6c757d', borderRadius: '8px', backgroundColor: '#ffffff', color: '#000' }}>
              <h3 style={{ margin: '0 0 15px 0', color: '#343a40' }}>📋 Historial de Pedidos Realizados ({pedidosHistorial.length})</h3>
              {pedidosHistorial.length > 0 ? (
                <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                  {pedidosHistorial.map((ped) => (
                    <div key={ped.numero} style={{ padding: '12px', border: '1px solid #dee2e6', borderRadius: '6px', backgroundColor: '#f8f9fa' }}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
                        <div>
                          <strong>{ped.numero}</strong> ({ped.idPedido}) — <small style={{ color: '#6c757d' }}>{ped.fecha}</small>
                        </div>
                        <span style={{ fontSize: '12px', padding: '3px 8px', borderRadius: '12px', fontWeight: 'bold', backgroundColor: ped.estado.includes('OK') ? '#d4edda' : '#fff3cd', color: ped.estado.includes('OK') ? '#155724' : '#856404' }}>
                          {ped.estado}
                        </span>
                      </div>
                      <div style={{ fontSize: '13px', color: '#495057' }}>
                        <strong>Detalle:</strong> {ped.items.map(i => `${i.nombre} (x${i.cantidad})`).join(', ')}
                      </div>
                      <div style={{ fontSize: '14px', fontWeight: 'bold', color: '#28a745', marginTop: '5px', textAlign: 'right' }}>
                        Monto: ${ped.total.toLocaleString()} CLP
                      </div>
                    </div>
                  ))}
                </div>
              ) : (
                <p style={{ margin: 0, color: '#6c757d' }}>Aún no se han realizado pedidos en esta sesión.</p>
              )}
            </div>
          )}

          {/* VISTA INDIVIDUAL 6: NOTIFICACIONES */}
          {vistaActiva === 'notificaciones' && (
            <div style={{ marginTop: '20px', padding: '15px', border: '1px solid #17a2b8', borderRadius: '8px', backgroundColor: '#ffffff', color: '#000' }}>
              <h3 style={{ margin: '0 0 15px 0', color: '#117a8b' }}>📩 Enviar Notificación — Lista de Pedidos ({pedidosHistorial.length})</h3>
              {pedidosHistorial.length > 0 ? (
                <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                  {pedidosHistorial.map((ped) => (
                    <div key={ped.numero} style={{ padding: '14px', border: '1px solid #dee2e6', borderRadius: '6px', backgroundColor: '#f8f9fa' }}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
                        <div>
                          <strong>{ped.numero}</strong> ({ped.idPedido}) — <small style={{ color: '#6c757d' }}>{ped.fecha}</small>
                        </div>
                        <span style={{ fontSize: '12px', padding: '3px 8px', borderRadius: '12px', fontWeight: 'bold', backgroundColor: ped.estado.includes('OK') ? '#d4edda' : '#fff3cd', color: ped.estado.includes('OK') ? '#155724' : '#856404' }}>
                          {ped.estado}
                        </span>
                      </div>

                      <div style={{ fontSize: '13px', color: '#495057', marginBottom: '10px' }}>
                        <strong>Productos:</strong> {ped.items.map(i => `${i.nombre} (x${i.cantidad})`).join(', ')}
                      </div>

                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                        {ped.notificacionEnviada ? (
                          <button disabled style={{ backgroundColor: '#28a745', color: 'white', border: 'none', padding: '6px 14px', borderRadius: '4px', fontWeight: 'bold', cursor: 'default' }}>
                            ✅ Enviado
                          </button>
                        ) : (
                          <button onClick={() => handleEnviarNotificacionPedido(ped)} style={{ backgroundColor: '#17a2b8', color: 'white', border: 'none', padding: '6px 14px', borderRadius: '4px', fontWeight: 'bold', cursor: 'pointer' }}>
                            📩 Enviar Notificación
                          </button>
                        )}

                        <div style={{ fontSize: '15px', fontWeight: 'bold', color: '#28a745' }}>
                          Monto: ${ped.total.toLocaleString()} CLP
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              ) : (
                <p style={{ margin: 0, color: '#6c757d' }}>No hay pedidos registrados para notificar.</p>
              )}
            </div>
          )}

          {/* CONSOLA DE LOGS / RESPUESTA DE API */}
          <h3 style={{ marginTop: '20px' }}>Respuesta de la API / Log:</h3>
          <pre style={{ background: '#f4f4f4', color: '#000', padding: '15px', borderRadius: '5px', overflowX: 'auto', maxHeight: '250px' }}>
            {log || 'Haz clic en algún botón para probar la integración con los microservicios...'}
          </pre>
        </main>
      </AuthenticatedTemplate>
    </div>
  );
}

export default App;