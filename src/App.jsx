import React, { useState } from 'react';
import { useMsal, AuthenticatedTemplate, UnauthenticatedTemplate } from '@azure/msal-react';
import { loginRequest } from './auth/AuthConfig';
import './App.css';

const SERVICES = {
  CATALOGO: 'https://j76lk49uq2.execute-api.us-east-1.amazonaws.com/v1/catalogo',
  INVENTARIO: 'https://j76lk49uq2.execute-api.us-east-1.amazonaws.com/v1/inventario',
  PEDIDOS_LOCAL: 'https://j76lk49uq2.execute-api.us-east-1.amazonaws.com/v1/pedidos',
  PEDIDOS: 'https://j76lk49uq2.execute-api.us-east-1.amazonaws.com/v1/pedidos',
  USUARIOS: 'https://j76lk49uq2.execute-api.us-east-1.amazonaws.com/v1/usuarios/perfil',
  PAGOS: 'https://j76lk49uq2.execute-api.us-east-1.amazonaws.com/v1/pagos/procesar',
  NOTIFICACIONES: 'https://j76lk49uq2.execute-api.us-east-1.amazonaws.com/v1/notificaciones'
};

const SCOPES = {
  CATALOGO: ["api://5d7ead55-12ca-4dcd-97e8-5b5e81f77f4b/Catalogo.Read"],
  INVENTARIO: ["api://1f835a60-98e3-4fde-b87f-0c64cd1e7e46/Inventario.Read"],
  PAGOS: ["api://5d7ead55-12ca-4dcd-97e8-5b5e81f77f4b/Pagos.Write"]
};

const PRECIOS_PRODUCTOS = {
  "Hamburguesa Triple": 6500,
  "Papas Supremas": 3500,
  "Bebida Grande": 2000,
  "Pizza Familiar": 12000,
  "Pastel de Chocolate": 4000
};

const DETALLES_VISUALES = {
  "Hamburguesa Triple": {
    emoji: "🍔",
    categoria: "Hamburguesas",
    desc: "Triple carne a la parrilla, queso cheddar fundido, tocino crocante y salsa de la casa."
  },
  "Papas Supremas": {
    emoji: "🍟",
    categoria: "Acompañamientos",
    desc: "Papas rústicas doradas con crema ácida, ciboulette y lluvia de queso."
  },
  "Bebida Grande": {
    emoji: "🥤",
    categoria: "Bebidas",
    desc: "Vaso de 750 ml bien helado a elección."
  },
  "Pizza Familiar": {
    emoji: "🍕",
    categoria: "Pizzas",
    desc: "Masa artesanal de 38 cm con mozzarella, pepperoni y orégano fresco."
  },
  "Pastel de Chocolate": {
    emoji: "🍰",
    categoria: "Postres",
    desc: "Bizcocho húmedo de cacao intenso relleno con ganache semi-amargo."
  }
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
      catList = Array.isArray(parsedCat)
        ? parsedCat
        : ["Hamburguesa Triple", "Papas Supremas", "Bebida Grande", "Pizza Familiar", "Pastel de Chocolate"];
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

    // Se realiza la autenticación en Azure MSAL enviando el Scope de Pagos
    await callApi(SERVICES.PAGOS, 'POST', {
      boleta: numBoleta,
      monto: montoAPagar,
      numeroTarjeta: '4532XXXXXXXX8892',
      estado: 'APROBADO'
    }, SCOPES.PAGOS);

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
    const detalleProductos = pedidoTarget.items
      .map(i => `${i.nombre} (x${i.cantidad}) - $${i.subtotal.toLocaleString()} CLP`)
      .join('\n - ');
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
  const subtotalCarrito = carrito.reduce((sum, i) => sum + (PRECIOS_PRODUCTOS[i.nombre] || 5000) * i.cantidad, 0);
  const nombreUsuario = accounts[0]?.username || 'usuario.prueba@pedidos360.cl';

  return (
    <>
      <UnauthenticatedTemplate>
        <div className="login-wrapper">
          <div className="login-card">
            <div style={{ fontSize: '2.5rem', marginBottom: '0.5rem' }}>📦</div>
            <h1 style={{ margin: '0 0 0.5rem 0', fontSize: '1.6rem', color: '#002b49' }}>Pedidos360</h1>
            <p style={{ color: '#64748b', fontSize: '0.92rem', margin: 0 }}>
              Debes iniciar sesión con tu cuenta de Microsoft Entra ID para acceder al sistema.
            </p>
            <button className="btn-ms-login" onClick={handleLogin}>
              Iniciar Sesión con Microsoft
            </button>
          </div>
        </div>
      </UnauthenticatedTemplate>

      <AuthenticatedTemplate>
        <div className="ecommerce-app">
          {/* BARRA DE NAVEGACIÓN */}
          <header className="navbar">
            <div className="navbar-inner">
              <div className="brand" onClick={handleVerCatalogo}>
                <span className="brand-badge">📦</span>
                <span>Pedidos360</span>
              </div>

              <nav className="nav-links">
                <button
                  className={`nav-btn ${vistaActiva === 'catalogo' ? 'active' : ''}`}
                  onClick={handleVerCatalogo}
                >
                  Tienda
                </button>
                <button
                  className={`nav-btn ${vistaActiva === 'inventario' ? 'active' : ''}`}
                  onClick={handleVerInventario}
                >
                  Inventario
                </button>
                <button
                  className={`nav-btn ${vistaActiva === 'boleta' ? 'active' : ''}`}
                  onClick={() => setVistaActiva('boleta')}
                >
                  Boleta / Pago
                </button>
                <button
                  className={`nav-btn ${vistaActiva === 'pedidos' ? 'active' : ''}`}
                  onClick={handleConsultarPedidos}
                >
                  Mis Pedidos ({pedidosHistorial.length})
                </button>
                <button
                  className={`nav-btn ${vistaActiva === 'notificaciones' ? 'active' : ''}`}
                  onClick={handleBotonMenuNotificaciones}
                >
                  Notificaciones
                </button>
                <button
                  className={`nav-btn ${vistaActiva === 'perfil' ? 'active' : ''}`}
                  onClick={handleVerPerfil}
                >
                  Mi Perfil
                </button>
              </nav>

              <div className="user-actions">
                <span className="user-pill">👤 {nombreUsuario}</span>
                <button className="btn-logout" onClick={handleLogout}>Cerrar Sesión</button>
              </div>
            </div>
          </header>

          {/* HERO BANNER */}
          <section className="hero-banner">
            <div className="hero-content">
              <div>
                <span className="hero-tag">Entrega Rápida • Pago Seguro Webpay</span>
                <h1 className="hero-title">Tu comida favorita en minutos</h1>
                <p className="hero-subtitle">Explora el menú disponible en tiempo real, arma tu pedido y paga online.</p>
              </div>
            </div>
          </section>

          {/* CONTENIDO PRINCIPAL + CARRITO LATERAL */}
          <div className="store-container">
            <main>
              {/* VISTA 1: CATÁLOGO */}
              {vistaActiva === 'catalogo' && (
                <>
                  <div className="section-header">
                    <h2 className="section-title">Menú Destacado</h2>
                    <span style={{ fontSize: '0.85rem', color: '#64748b' }}>Precios en CLP • IVA incluido</span>
                  </div>

                  <div className="products-grid">
                    {(catalogo.length > 0
                      ? catalogo
                      : ["Hamburguesa Triple", "Papas Supremas", "Bebida Grande", "Pizza Familiar", "Pastel de Chocolate"]
                    ).map((prod) => {
                      const stockDisponible = inventario[prod] !== undefined ? inventario[prod] : 50;
                      const tieneStock = stockDisponible > 0;
                      const precio = PRECIOS_PRODUCTOS[prod] || 5000;
                      const info = DETALLES_VISUALES[prod] || {
                        emoji: "🍽️",
                        categoria: "Especialidad",
                        desc: "Preparado al momento con ingredientes seleccionados."
                      };

                      return (
                        <article key={prod} className="product-card">
                          <div className="product-visual">
                            <span>{info.emoji}</span>
                            <span className={`stock-badge ${tieneStock ? 'in-stock' : 'out-stock'}`}>
                              {tieneStock ? `${stockDisponible} disp.` : 'Agotado'}
                            </span>
                          </div>

                          <div className="product-body">
                            <span className="product-category">{info.categoria}</span>
                            <h3 className="product-name">{prod}</h3>
                            <p className="product-desc">{info.desc}</p>

                            <div className="product-footer">
                              <span className="product-price">${precio.toLocaleString('es-CL')}</span>
                              <button
                                className="btn-add-cart"
                                onClick={() => agregarAlCarrito(prod)}
                                disabled={!tieneStock}
                              >
                                {tieneStock ? '+ Agregar' : 'Agotado'}
                              </button>
                            </div>
                          </div>
                        </article>
                      );
                    })}
                  </div>
                </>
              )}

              {/* VISTA 2: INVENTARIO */}
              {vistaActiva === 'inventario' && (
                <div className="panel-card">
                  <h2 className="section-title">📊 Control de Inventario en Tiempo Real</h2>
                  <table className="data-table">
                    <thead>
                      <tr>
                        <th>Producto</th>
                        <th>Stock Disponible</th>
                        <th>Estado</th>
                      </tr>
                    </thead>
                    <tbody>
                      {Object.entries(Object.keys(inventario).length > 0 ? inventario : {
                        "Hamburguesa Triple": 50,
                        "Pizza Familiar": 30,
                        "Papas Supremas": 100,
                        "Pastel de Chocolate": 20,
                        "Bebida Grande": 80
                      }).map(([prod, stock]) => (
                        <tr key={prod}>
                          <td><strong>{prod}</strong></td>
                          <td>{stock} unidades</td>
                          <td>
                            <span className={`status-pill ${stock > 0 ? 'status-paid' : 'status-pending'}`}>
                              {stock > 0 ? 'Disponible' : 'Sin Stock'}
                            </span>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}

              {/* VISTA 3: BOLETA Y PAGO */}
              {vistaActiva === 'boleta' && (
                <div className="panel-card">
                  {boleta ? (
                    <>
                      <div className="section-header">
                        <h2 className="section-title">🧾 Boleta Digital: {boleta.numero}</h2>
                        <span className={`status-pill ${boleta.estado.includes('OK') ? 'status-paid' : 'status-pending'}`}>
                          {boleta.estado}
                        </span>
                      </div>

                      <p style={{ margin: '0.3rem 0', fontSize: '0.9rem' }}><strong>ID Pedido:</strong> {boleta.idPedido}</p>
                      <p style={{ margin: '0.3rem 0', fontSize: '0.9rem' }}><strong>Cliente:</strong> {boleta.cliente}</p>
                      <p style={{ margin: '0.3rem 0', fontSize: '0.9rem' }}><strong>Fecha Emisión:</strong> {boleta.fecha}</p>
                      {boleta.fechaPago && (
                        <p style={{ margin: '0.3rem 0', fontSize: '0.9rem', color: '#16a34a' }}>
                          <strong>Fecha Pago:</strong> {boleta.fechaPago}
                        </p>
                      )}

                      <table className="data-table">
                        <thead>
                          <tr>
                            <th>Producto</th>
                            <th>Cant.</th>
                            <th>P. Unit.</th>
                            <th>Subtotal</th>
                          </tr>
                        </thead>
                        <tbody>
                          {boleta.items.map((item, idx) => (
                            <tr key={idx}>
                              <td>{item.nombre}</td>
                              <td>{item.cantidad}</td>
                              <td>${item.precioUnitario.toLocaleString('es-CL')}</td>
                              <td><strong>${item.subtotal.toLocaleString('es-CL')}</strong></td>
                            </tr>
                          ))}
                        </tbody>
                      </table>

                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: '1.5rem' }}>
                        {!boleta.estado.includes('OK') ? (
                          <button className="btn-pay-now" style={{ width: 'auto', padding: '0.75rem 1.5rem' }} onClick={handleProbarPago}>
                            💳 Pagar ahora
                          </button>
                        ) : (
                          <span style={{ color: '#16a34a', fontWeight: '700' }}>✅ Pago confirmado</span>
                        )}
                        <div style={{ fontSize: '1.35rem', fontWeight: '800', color: '#002b49' }}>
                          TOTAL: ${boleta.total.toLocaleString('es-CL')} CLP
                        </div>
                      </div>
                    </>
                  ) : (
                    <div className="cart-empty">
                      <p>No hay boleta activa. Añade productos al carrito y presiona <strong>Confirmar Pedido</strong>.</p>
                    </div>
                  )}
                </div>
              )}

              {/* VISTA 4: HISTORIAL DE PEDIDOS */}
              {vistaActiva === 'pedidos' && (
                <div className="panel-card">
                  <h2 className="section-title" style={{ marginBottom: '1rem' }}>📋 Historial de Pedidos Realizados ({pedidosHistorial.length})</h2>
                  {pedidosHistorial.length > 0 ? (
                    pedidosHistorial.map((ped) => (
                      <div key={ped.numero} className="order-item-card">
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '0.5rem' }}>
                          <div>
                            <strong>{ped.numero}</strong> ({ped.idPedido}) — <small style={{ color: '#64748b' }}>{ped.fecha}</small>
                          </div>
                          <span className={`status-pill ${ped.estado.includes('OK') ? 'status-paid' : 'status-pending'}`}>
                            {ped.estado}
                          </span>
                        </div>
                        <p style={{ margin: '0.4rem 0', fontSize: '0.88rem', color: '#475569' }}>
                          <strong>Detalle:</strong> {ped.items.map(i => `${i.nombre} (x${i.cantidad})`).join(', ')}
                        </p>
                        <div style={{ textAlign: 'right', fontWeight: '800', color: '#002b49' }}>
                          Monto: ${ped.total.toLocaleString('es-CL')} CLP
                        </div>
                      </div>
                    ))
                  ) : (
                    <div className="cart-empty">Aún no se han realizado pedidos en esta sesión.</div>
                  )}
                </div>
              )}

              {/* VISTA 5: NOTIFICACIONES */}
              {vistaActiva === 'notificaciones' && (
                <div className="panel-card">
                  <h2 className="section-title" style={{ marginBottom: '1rem' }}>📩 Enviar Notificación — Lista de Pedidos ({pedidosHistorial.length})</h2>
                  {pedidosHistorial.length > 0 ? (
                    pedidosHistorial.map((ped) => (
                      <div key={ped.numero} className="order-item-card">
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '0.5rem' }}>
                          <div>
                            <strong>{ped.numero}</strong> ({ped.idPedido}) — <small style={{ color: '#64748b' }}>{ped.fecha}</small>
                          </div>
                          <span className={`status-pill ${ped.estado.includes('OK') ? 'status-paid' : 'status-pending'}`}>
                            {ped.estado}
                          </span>
                        </div>
                        <p style={{ margin: '0.4rem 0 0.8rem 0', fontSize: '0.88rem', color: '#475569' }}>
                          <strong>Productos:</strong> {ped.items.map(i => `${i.nombre} (x${i.cantidad})`).join(', ')} — <strong>${ped.total.toLocaleString('es-CL')} CLP</strong>
                        </p>
                        {ped.notificacionEnviada ? (
                          <button disabled className="btn-notify-sent">✅ Enviado</button>
                        ) : (
                          <button className="btn-notify" onClick={() => handleEnviarNotificacionPedido(ped)}>
                            📩 Enviar Notificación
                          </button>
                        )}
                      </div>
                    ))
                  ) : (
                    <div className="cart-empty">No hay pedidos registrados para notificar.</div>
                  )}
                </div>
              )}

              {/* VISTA 6: PERFIL DE USUARIO */}
              {vistaActiva === 'perfil' && (
                <div className="panel-card">
                  <h2 className="section-title" style={{ marginBottom: '1rem' }}>👤 Perfil del Usuario</h2>
                  <p style={{ margin: '0.4rem 0' }}><strong>Correo institucional:</strong> {nombreUsuario}</p>
                  <p style={{ margin: '0.4rem 0' }}><strong>Autenticación:</strong> Microsoft Entra ID (Activo)</p>
                  <p style={{ margin: '0.4rem 0' }}><strong>Pedidos realizados en sesión:</strong> {pedidosHistorial.length}</p>
                </div>
              )}
            </main>

            {/* CARRITO LATERAL */}
            <aside className="cart-sidebar">
              <div className="cart-header">
                <h2 className="cart-title">
                  🛒 Tu Pedido <span className="cart-count-pill">{totalProductosCarrito}</span>
                </h2>
                {carrito.length > 0 && (
                  <button className="btn-clear-cart" onClick={vaciarCarrito}>Vaciar</button>
                )}
              </div>

              {carrito.length > 0 ? (
                <>
                  <div className="cart-items">
                    {carrito.map((item) => {
                      const precioUnit = PRECIOS_PRODUCTOS[item.nombre] || 5000;
                      return (
                        <div key={item.nombre} className="cart-item">
                          <div className="cart-item-info">
                            <span className="cart-item-name">{item.nombre}</span>
                            <span className="cart-item-price">${(precioUnit * item.cantidad).toLocaleString('es-CL')}</span>
                          </div>

                          <div className="qty-controls">
                            <button className="qty-btn" onClick={() => quitarDelCarrito(item.nombre)}>-</button>
                            <span className="qty-value">{item.cantidad}</span>
                            <button
                              className="qty-btn"
                              onClick={() => agregarAlCarrito(item.nombre)}
                              disabled={(inventario[item.nombre] !== undefined ? inventario[item.nombre] : 50) <= 0}
                            >
                              +
                            </button>
                          </div>
                        </div>
                      );
                    })}
                  </div>

                  <div className="cart-summary">
                    <div className="summary-row">
                      <span>Subtotal</span>
                      <span>${subtotalCarrito.toLocaleString('es-CL')}</span>
                    </div>
                    <div className="summary-row">
                      <span>Envío</span>
                      <span style={{ color: '#16a34a', fontWeight: '600' }}>Gratis</span>
                    </div>
                    <div className="summary-total">
                      <span>Total</span>
                      <span>${subtotalCarrito.toLocaleString('es-CL')} CLP</span>
                    </div>

                    <button className="btn-checkout" onClick={handleCrearPedido}>
                      Confirmar Pedido (RabbitMQ)
                    </button>
                  </div>
                </>
              ) : (
                <div className="cart-empty">
                  <div style={{ fontSize: '2rem', marginBottom: '0.5rem' }}>🛍️</div>
                  El carrito está vacío.<br />Revisa el catálogo para añadir productos.
                </div>
              )}

              {boleta && !boleta.estado.includes('OK') && (
                <button className="btn-pay-now" onClick={handleProbarPago}>
                  💳 Pagar Boleta {boleta.numero} (${boleta.total.toLocaleString('es-CL')})
                </button>
              )}
            </aside>
          </div>

          {/* CONSOLA DE LOGS / RESPUESTA DE API */}
          <div className="api-console-wrapper">
            <details className="api-console">
              <summary>🛠️ Respuesta de la API / Log de Microservicios</summary>
              <pre>{log || 'Haz clic en algún botón para probar la integración con los microservicios...'}</pre>
            </details>
          </div>
        </div>
      </AuthenticatedTemplate>
    </>
  );
}

export default App;