# CENTRO DE FOTOS PARA EJECUCION PDV V2

import json
import pandas as pd

# Variable con la ruta del archivo de Excel
excel_file = "Base principal.xlsx"
df = pd.read_excel(excel_file)

# Limpiar espacios al inicio y final en los nombres de las columnas
df.columns = df.columns.str.strip()

# Llenar valores nulos (NaN) con texto vacío para evitar errores
df = df.fillna("")

# Convertir cualquier columna de fecha a texto en formato YYYY-MM-DD
for col in df.columns:
    if pd.api.types.is_datetime64_any_dtype(df[col]):
        df[col] = df[col].dt.strftime('%Y-%m-%d')

# Calcular la fecha máxima (fecha de corte) de los datos si existe la columna Fecha
fecha_corte_str = "--/--/----"
if "Fecha" in df.columns:
    valid_dates = [str(d).strip() for d in df["Fecha"] if d and str(d).strip() != ""]
    if valid_dates:
        valid_dates.sort()
        max_date = valid_dates[-1]  # Última fecha
        parts = max_date.split("-")
        if len(parts) == 3:
            fecha_corte_str = f"{parts[2]}/{parts[1]}/{parts[0]}"
        else:
            fecha_corte_str = max_date

# Convertir el DataFrame a JSON y ordenar por fecha de manera descendente (más reciente primero)
records = df.to_dict(orient="records")
records.sort(key=lambda x: str(x.get("Fecha", "")), reverse=True)

data_json = json.dumps(records, ensure_ascii=False)

# Generar el HTML con la lógica solicitada
html_content = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>Reporte de Ejecución de Campo</title>
<!-- FontAwesome para iconos -->
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
<style>
:root{{
  --bg:#f4f7fb;--panel:#ffffff;--ink:#132238;--muted:#6b7b91;--line:#dce5ef;
  --navy:#10233f;--blue:#2264e5;--blue-hover:#1d4ed8;--cyan:#0ea5a8;--green:#20a46b;--amber:#f2a93b;--red:#e05252;
  --shadow:0 8px 28px rgba(31,54,88,.08);--radius:18px;
}}
*{{box-sizing:border-box}} 
body{{margin:0;background:var(--bg);color:var(--ink);font-family:Inter,Segoe UI,Roboto,Arial,sans-serif}}
.app{{display:flex;min-height:100vh}}
.sidebar{{width:260px;background:linear-gradient(180deg,#10233f 0%,#17385f 100%);color:white;padding:24px 18px;position:fixed;inset:0 auto 0 0;overflow:auto;z-index:20}}

/* LOGO CORPORATIVO EN SIDEBAR */
.brand{{display:flex;align-items:center;justify-content:center;margin-bottom:28px}}
.brand img.logo-corp{{max-width:190px;height:auto;display:block;margin:0 auto}}

/* BOTONES LATERALES */
.navbtn{{
  width:100%;
  display:flex;
  gap:14px;
  align-items:center;
  border:0;
  background:transparent;
  color:#cfdded;
  padding:14px 14px;
  border-radius:12px;
  margin:6px 0;
  text-align:left;
  cursor:pointer;
  font-weight:700;
  font-size:15px;
  transition: all 0.2s ease;
}}
.navbtn:hover, .navbtn.active{{
  background:var(--blue);
  color:#ffffff;
}}
.navbtn span:first-child{{
  font-size:22px;
  width:26px;
  text-align:center;
}}

.main{{margin-left:260px;width:calc(100% - 260px);padding:24px 28px 42px}}

/* HEADER SUPERIOR CON FECHA DE CORTE */
.top-header{{
  display:flex;
  align-items:center;
  justify-content:space-between;
  background:var(--panel);
  padding:16px 24px;
  border-radius:var(--radius);
  box-shadow:var(--shadow);
  border:1px solid #e8eef5;
  margin-bottom:20px;
  flex-wrap:wrap;
  gap:15px;
}}
.header-left{{
  display:flex;
  align-items:center;
  gap:16px;
}}
.header-left img.logo-corp{{
  max-height:45px;
  width:auto;
}}
.header-left h1{{
  margin:0;
  font-size:20px;
  font-weight:800;
  color:var(--navy);
  letter-spacing:-.3px;
}}
.header-right{{
  display:flex;
  align-items:center;
  gap:8px;
  background:#f1f5f9;
  padding:8px 14px;
  border-radius:10px;
  font-size:13px;
  color:var(--muted);
  font-weight:600;
}}
.header-right i{{
  color:var(--blue);
  font-size:16px;
}}
.header-right strong{{
  color:var(--ink);
}}

/* BARRA SUPERIOR CON TÍTULO */
.topbar{{display:flex;align-items:center;justify-content:space-between;gap:20px;margin-bottom:18px;flex-wrap:wrap}}
.title h2{{margin:0;font-size:26px;letter-spacing:-.4px}}
.topbar-logo img{{max-height:55px;width:auto;display:block}}

/* BARRA DE FILTROS */
.filterbar{{background:var(--panel);padding:14px 16px 12px;border-radius:var(--radius);box-shadow:var(--shadow);border:1px solid #e8eef5;margin-bottom:18px;position:sticky;top:10px;z-index:15}}
.filterbar-head{{display:flex;align-items:flex-start;justify-content:space-between;gap:14px;margin-bottom:12px}}
.filterbar-title h3{{margin:3px 0 4px;font-size:17px;line-height:1.05;color:#17304f}}
.filterbar-title p{{margin:0;color:#6d7b90;font-size:11px;max-width:560px}}
.filterbar-tools{{display:flex;align-items:center;gap:10px;flex-wrap:wrap;justify-content:flex-end}}
.toggle-btn{{min-width:140px;font-weight:800}}

.filterbar-body{{display:block}}
.filterbar.collapsed{{padding-bottom:10px}}
.filterbar.collapsed .filterbar-head{{margin-bottom:10px}}
.filterbar.collapsed .filterbar-body{{display:none}}
.filterbar.collapsed .active-filter-strip{{margin-top:0}}
.filterbar.collapsed .filterbar-title p{{display:none}}
.filterbar.collapsed .filterbar-title h3{{font-size:15px;margin-bottom:0}}

/* GRID DE FILTROS EN DOS LÍNEAS EXACTAS */
.filters{{
  display:grid;
  grid-template-columns: repeat(5, 1fr);
  gap:12px 14px;
}}
.filter-field{{min-width:0; width:100%; position:relative;}}
.filter-field label{{
  display:block;
  color:#64758b;
  font-size:10px;
  font-weight:800;
  text-transform:none;
  letter-spacing:.3px;
  margin:0 0 4px 2px;
}}

.date-range{{
  display:flex;
  gap:6px;
  width:100%;
}}
.date-range input{{
  flex:1;
  min-width:0;
}}

.filters select, .filters input{{
  width:100%;
  border:1px solid var(--line);
  background:#fbfdff;
  color:var(--ink);
  border-radius:10px;
  padding:8px 10px;
  font-size:11px;
  outline:none;
}}
.filters select:disabled{{background:#f1f5f9;color:#94a3b8;cursor:not-allowed;border-color:#e2e8f0}}
.filters select:focus,.filters input:focus{{border-color:#74a0f3;box-shadow:0 0 0 3px rgba(34,100,229,.09)}}

/* ESTILOS PARA DROPDOWNS CON BÚSQUEDA */
.searchable-select-wrapper {{
  position: relative;
  width: 100%;
}}
.searchable-input {{
  width: 100%;
  border: 1px solid var(--line);
  background: #fbfdff;
  color: var(--ink);
  border-radius: 10px;
  padding: 8px 24px 8px 10px;
  font-size: 11px;
  outline: none;
}}
.searchable-input:disabled {{
  background: #f1f5f9;
  color: #94a3b8;
  cursor: not-allowed;
  border-color: #e2e8f0;
}}
.searchable-dropdown {{
  position: absolute;
  top: 100%;
  left: 0;
  width: 100%;
  max-height: 200px;
  overflow-y: auto;
  background: #ffffff;
  border: 1px solid var(--line);
  border-radius: 8px;
  box-shadow: 0 6px 16px rgba(0,0,0,0.1);
  z-index: 50;
  display: none;
  margin-top: 2px;
}}
.searchable-dropdown.active {{
  display: block;
}}
.searchable-option {{
  padding: 8px 10px;
  font-size: 11px;
  cursor: pointer;
  color: var(--ink);
  border-bottom: 1px solid #f1f5f9;
}}
.searchable-option:hover {{
  background: #eef4ff;
  color: var(--blue);
}}

/* BOTONES */
.filter-actions{{display:flex;align-items:center;gap:8px;margin-top:14px;flex-wrap:wrap}}
.btn{{
  border:0;
  border-radius:10px;
  padding:9px 14px;
  font-size:11px;
  font-weight:750;
  cursor:pointer;
  transition: background-color 0.2s ease, color 0.2s ease;
}}
.btn.primary{{background:var(--blue);color:white}}
.btn.ghost{{background:#edf3fa;color:#2a4365}}

.btn:hover{{
  background:var(--blue) !important;
  color:#ffffff !important;
}}

.active-filter-strip{{display:flex;align-items:center;gap:8px;flex-wrap:wrap;border-top:1px solid #ebf0f6;padding-top:10px;margin-top:10px}}
.active-filter-strip .caption{{font-size:9px;font-weight:850;color:#7b8da4;text-transform:uppercase;letter-spacing:.5px;margin-right:2px}}
.filter-chip{{display:inline-flex;align-items:center;gap:5px;background:#eef4ff;color:#2456a8;border:1px solid #dbe8ff;border-radius:999px;padding:4px 8px;font-size:9px;font-weight:750}}

/* GALERÍA DE TARJETAS */
.photo-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(270px,1fr));gap:20px;margin-top:16px}}

.card{{
  background:var(--panel);
  border-radius:20px;
  overflow:hidden;
  box-shadow:0 10px 25px rgba(0,0,0,0.05);
  border:1px solid #eef2f7;
  display:flex;
  flex-direction:column;
  transition:transform 0.2s ease, box-shadow 0.2s ease;
  cursor:pointer;
}}
.card:hover{{
  transform:translateY(-4px);
  box-shadow:0 14px 30px rgba(0,0,0,0.08);
}}

.card-img-wrapper{{
  position:relative;
  width:100%;
  height:210px;
  background:#e2e8f0;
  overflow:hidden;
  display:flex;
}}
.card-img-wrapper img{{
  width:100%;
  height:100%;
  object-fit:cover;
}}
.card-img-wrapper.dual img{{
  width:50%;
  border-right:1px solid #ffffff;
}}
.card-img-wrapper.dual img:last-child{{
  border-right:none;
}}

.card-badge{{
  position:absolute;
  top:12px;
  left:12px;
  background:rgba(255,255,255,0.92);
  backdrop-filter:blur(4px);
  color:#1e293b;
  font-size:10px;
  font-weight:800;
  padding:4px 10px;
  border-radius:20px;
  text-transform:uppercase;
  letter-spacing:0.5px;
  box-shadow:0 2px 6px rgba(0,0,0,0.12);
}}

.card-body{{
  padding:16px;
  display:flex;
  flex-direction:column;
  gap:8px;
}}
.card-title{{
  margin:0 0 4px 0;
  font-size:16px;
  font-weight:800;
  color:#0f172a;
  line-height:1.2;
}}
.card-info-grid{{
  display:grid;
  grid-template-columns: 1fr 1fr;
  gap:8px;
  margin-top:4px;
}}
.card-info-item{{
  display:flex;
  align-items:center;
  gap:6px;
  color:#64748b;
  font-size:12px;
  font-weight:600;
  white-space:nowrap;
  overflow:hidden;
  text-overflow:ellipsis;
}}
.card-info-item .icon{{
  font-size:14px;
  flex-shrink:0;
}}
.card-info-item.full-width{{
  grid-column: span 2;
}}

/* PAGINACIÓN */
.pagination-container{{
  display:flex;
  align-items:center;
  justify-content:space-between;
  margin-top:24px;
  padding:12px 16px;
  background:var(--panel);
  border-radius:var(--radius);
  border:1px solid #e8eef5;
  box-shadow:var(--shadow);
}}
.pagination-info{{font-size:12px;font-weight:700;color:var(--muted)}}
.pagination-buttons{{display:flex;gap:8px}}

/* LIGHTBOX MODAL */
.modal-overlay{{
  position:fixed;
  inset:0;
  background:rgba(10, 18, 30, 0.88);
  backdrop-filter:blur(6px);
  z-index:100;
  display:none;
  align-items:center;
  justify-content:center;
  padding:15px;
}}
.modal-overlay.active{{
  display:flex;
}}
.modal-content{{
  background:#ffffff;
  border-radius:20px;
  width:95vw;
  height:92vh;
  display:flex;
  overflow:hidden;
  box-shadow:0 25px 50px -12px rgba(0, 0, 0, 0.4);
  position:relative;
}}
.modal-close{{
  position:absolute;
  top:20px;
  right:20px;
  background:transparent;
  border:0;
  font-size:24px;
  cursor:pointer;
  z-index:10;
  display:flex;
  align-items:center;
  justify-content:center;
  color:#0f172a;
  font-weight:bold;
  transition:transform 0.2s;
}}
.modal-close:hover{{transform:scale(1.2); color:#e05252;}}

.modal-viewer{{
  flex:2.8;
  background:#090d16;
  display:flex;
  align-items:center;
  justify-content:center;
  padding:30px;
  position:relative;
}}
.modal-main-img{{
  max-width:100%;
  max-height:100%;
  object-fit:contain;
  border-radius:8px;
}}

.modal-sidebar{{
  flex:1.2;
  min-width:340px;
  max-width:420px;
  padding:32px 28px;
  display:flex;
  flex-direction:column;
  justify-content:space-between;
  background:#ffffff;
  border-left:1px solid #e2e8f0;
  overflow-y:auto;
}}
.modal-sidebar-top h3{{
  margin:0 0 20px 0;
  font-size:20px;
  font-weight:800;
  color:#0f172a;
  border-bottom:1px solid #e2e8f0;
  padding-bottom:12px;
}}
.modal-info-list{{
  display:grid;
  grid-template-columns: 1fr 1fr;
  gap:16px 12px;
}}
.modal-info-item{{
  display:flex;
  flex-direction:column;
  gap:2px;
}}
.modal-info-item label{{
  display:block;
  font-size:10px;
  color:#94a3b8;
  font-weight:800;
  text-transform:uppercase;
  letter-spacing:0.5px;
}}
.modal-info-item span{{
  font-size:13px;
  font-weight:700;
  color:#1e293b;
  word-break:break-word;
}}

/* SECCIÓN INFERIOR: FOTOS DEL REGISTRO */
.modal-photos-section{{
  border-top:1px solid #e2e8f0;
  padding-top:18px;
  margin-top:20px;
}}
.modal-photos-section label{{
  display:block;
  font-size:11px;
  font-weight:850;
  color:#64748b;
  text-transform:uppercase;
  letter-spacing:0.5px;
  margin-bottom:12px;
}}
.modal-thumbnails{{
  display:flex;
  gap:12px;
  flex-wrap:wrap;
}}
.thumb-item{{
  display:flex;
  flex-direction:column;
  align-items:center;
  gap:4px;
}}
.thumb-btn{{
  background:transparent;
  border:2px solid transparent;
  border-radius:8px;
  padding:2px;
  cursor:pointer;
  transition:all 0.2s;
  overflow:hidden;
}}
.thumb-btn.active{{
  border-color:var(--blue);
  box-shadow:0 0 0 2px rgba(34,100,229,0.2);
}}
.thumb-btn img{{
  width:55px;
  height:55px;
  object-fit:cover;
  border-radius:6px;
  display:block;
}}
.thumb-label{{
  font-size:10px;
  font-weight:700;
  color:#64748b;
  text-transform:uppercase;
}}

@media(max-width:1300px){{
  .filters{{grid-template-columns: repeat(3, 1fr)}}
}}
@media(max-width:850px){{
  .sidebar{{position:relative;width:100%}}
  .app{{display:block}}
  .main{{margin-left:0;width:100%;padding:16px}}
  .filters{{grid-template-columns: 1fr 1fr}}
  .modal-content{{flex-direction:column;width:98vw;height:96vh;overflow-y:auto;}}
  .modal-viewer{{height:50vh;padding:15px;}}
  .modal-sidebar{{min-width:100%;max-width:100%;}}
}}
</style>
</head>
<body>

<div class="app">
<aside class="sidebar">
  <div class="brand">
    <img src="https://mpw-pwt.github.io/Ejecuciones/Manpower%20Group.png" alt="ManpowerGroup Logo" class="logo-corp" id="corpLogo"/>
  </div>
  
  <button class="navbtn active" data-exec="Surtido"><span>📦</span><span>Surtido</span></button>
  <button class="navbtn" data-exec="Exhibición adicional"><span>🖼️</span><span>Exhibición adicional</span></button>
  <button class="navbtn" data-exec="Exhibición paga / convenio"><span>🤝</span><span>Exhibición paga / convenio</span></button>
  <button class="navbtn" data-exec="POP"><span>🚩</span><span>POP</span></button>
  <button class="navbtn" data-exec="Impulso"><span>🚀</span><span>Impulso</span></button>
  <button class="navbtn" data-exec="Descuentos"><span>🏷️</span><span>Descuentos</span></button>
</aside>

<main class="main">
  <!-- ENCABEZADO SUPERIOR CON FECHA DE CORTE -->
  <header class="top-header">
    <div class="header-left">
      <img src="https://mpw-pwt.github.io/Ejecuciones/Manpower%20Group.png" alt="Logo Corporativo" class="logo-corp">
      <h1 id="tituloHeader" style="font-size: 22px; font-weight: 800; color: var(--navy); margin:0;">CENTRO DE FOTOS PARA EJECUCIÓN PDV</h1>
    </div>
    <div class="header-right">
      <i class="fa-regular fa-calendar-check"></i>
      <span>Fecha de corte: <strong id="lblFechaCorte">{fecha_corte_str}</strong></span>
    </div>
  </header>

  <div class="topbar">
    <div class="title">
      <h2 id="pageTitle">Surtido</h2>
    </div>
    <div class="topbar-logo">
      <img src="https://mpw-pwt.github.io/Ejecuciones/Logo%20verde-02.png" alt="Logo Secundaria" />
    </div>
  </div>

  <section class="filterbar" id="filterBar">
    <div class="filterbar-head">
      <div class="filterbar-title">
        <h3>Filtros del tablero</h3>
        <p>Aplica los filtros que necesites y luego ocúltalos para ganar espacio en la visualización.</p>
      </div>
      <div class="filterbar-tools">
        <button class="btn ghost toggle-btn" id="toggleFiltersBtn" aria-expanded="true">Ocultar filtros</button>
      </div>
    </div>

    <div class="filterbar-body" id="filterBody">
      <div class="filters">
        <!-- Primera línea: Rango de fechas, distrito, canal, cadena y ciudad -->
        <div class="filter-field">
          <label for="fDateFrom">Rango de fechas</label>
          <div class="date-range">
            <input type="date" id="fDateFrom" aria-label="Fecha desde"/>
            <input type="date" id="fDateTo" aria-label="Fecha hasta"/>
          </div>
        </div>
        <div class="filter-field">
          <label>Distrito</label>
          <div class="searchable-select-wrapper" data-filter="fDistrict">
            <input type="text" class="searchable-input" id="fDistrict" placeholder="Todos los distritos" autocomplete="off"/>
            <div class="searchable-dropdown" id="drop_fDistrict"></div>
          </div>
        </div>
        <div class="filter-field">
          <label>Canal</label>
          <div class="searchable-select-wrapper" data-filter="fChannel">
            <input type="text" class="searchable-input" id="fChannel" placeholder="Todos los canales" autocomplete="off"/>
            <div class="searchable-dropdown" id="drop_fChannel"></div>
          </div>
        </div>
        <div class="filter-field">
          <label>Cadena</label>
          <div class="searchable-select-wrapper" data-filter="fChain">
            <input type="text" class="searchable-input" id="fChain" placeholder="Todas las cadenas" autocomplete="off"/>
            <div class="searchable-dropdown" id="drop_fChain"></div>
          </div>
        </div>
        <div class="filter-field">
          <label>Ciudad</label>
          <div class="searchable-select-wrapper" data-filter="fCity">
            <input type="text" class="searchable-input" id="fCity" placeholder="Todas las ciudades" autocomplete="off"/>
            <div class="searchable-dropdown" id="drop_fCity"></div>
          </div>
        </div>

        <!-- Segunda línea: Punto de venta, categoría y tipo de exhibición -->
        <div class="filter-field" style="grid-column: span 2;">
          <label>Punto de venta</label>
          <div class="searchable-select-wrapper" data-filter="fPuntoVenta">
            <input type="text" class="searchable-input" id="fPuntoVenta" placeholder="Todos los puntos" autocomplete="off"/>
            <div class="searchable-dropdown" id="drop_fPuntoVenta"></div>
          </div>
        </div>
        <div class="filter-field" style="grid-column: span 1.5;">
          <label>Categoría</label>
          <div class="searchable-select-wrapper" data-filter="fCategory">
            <input type="text" class="searchable-input" id="fCategory" placeholder="Todas las categorías" autocomplete="off" disabled/>
            <div class="searchable-dropdown" id="drop_fCategory"></div>
          </div>
        </div>
        <div class="filter-field" style="grid-column: span 1.5;">
          <label>Tipo de exhibición</label>
          <div class="searchable-select-wrapper" data-filter="fExhibType">
            <input type="text" class="searchable-input" id="fExhibType" placeholder="Todas las exhibiciones" autocomplete="off" disabled/>
            <div class="searchable-dropdown" id="drop_fExhibType"></div>
          </div>
        </div>
      </div>

      <div class="filter-actions">
        <button class="btn ghost" id="clearBtn">Restablecer</button>
        <button class="btn ghost" id="exportBtn">Exportar a excel</button>
      </div>
    </div>

    <div class="active-filter-strip" id="activeFilters">
      <span class="caption">Filtros activos:</span>
    </div>
  </section>

  <div class="photo-grid" id="photoGrid"></div>
  
  <!-- Contenedor de Paginación -->
  <div class="pagination-container" id="paginationContainer" style="display:none;">
    <div class="pagination-info" id="paginationInfo">Mostrando 0 - 0 de 0 registros</div>
    <div class="pagination-buttons">
      <button class="btn ghost" id="prevPageBtn" onclick="changePage(-1)">Anterior</button>
      <button class="btn primary" id="nextPageBtn" onclick="changePage(1)">Siguiente</button>
    </div>
  </div>
</main>
</div>

<!-- LIGHTBOX MODAL -->
<div class="modal-overlay" id="lightboxModal">
  <div class="modal-content">
    <button class="modal-close" id="closeModalBtn">✕</button>
    <div class="modal-viewer">
      <img id="modalImg" class="modal-main-img" src="" alt="Vista ampliada"/>
    </div>
    <div class="modal-sidebar">
      <div class="modal-sidebar-top">
        <h3 id="mPuntoVenta">-</h3>
        <div class="modal-info-list">
          <div class="modal-info-item">
            <label>Nombre Completo</label>
            <span id="mNombreUsuario">-</span>
          </div>
          <div class="modal-info-item">
            <label>Ciudad</label>
            <span id="mCiudad">-</span>
          </div>
          <div class="modal-info-item">
            <label>Fecha</label>
            <span id="mFecha">-</span>
          </div>
          <div class="modal-info-item" id="mTipoContainer" style="display:none;">
            <label id="mTipoLabel">Tipo</label>
            <span id="mTipo">-</span>
          </div>
        </div>
      </div>

      <div class="modal-photos-section">
        <label id="mPhotosTitle">FOTOS DEL REGISTRO (0)</label>
        <div class="modal-thumbnails" id="modalThumbnails"></div>
      </div>
    </div>
  </div>
</div>

<script>
const rawData = {data_json};

let currentExec = "Surtido";
let currentPage = 1;
const pageSize = 20;

function initDateConstraints() {{
  const dates = rawData.map(item => item["Fecha"]).filter(d => d && d.trim() !== "");
  if (dates.length > 0) {{
    dates.sort();
    const minDate = dates[0];
    const maxDate = dates[dates.length - 1];
    
    const dateFrom = document.getElementById("fDateFrom");
    const dateTo = document.getElementById("fDateTo");
    
    dateFrom.min = minDate;
    dateFrom.max = maxDate;
    dateTo.min = minDate;
    dateTo.max = maxDate;
  }}
}}

function getUserName(item) {{
  if (!item) return "Sin usuario";
  
  const keys = Object.keys(item);
  for (let key of keys) {{
    const cleanKey = key.toLowerCase().normalize("NFD").replace(/[\\u0300-\\u036f]/g, "").trim();
    if (
      cleanKey === "nombre completo" ||
      cleanKey === "nombrecompleto" ||
      cleanKey === "nombre usuario" ||
      cleanKey === "nombre" ||
      cleanKey === "usuario" ||
      cleanKey === "promotor"
    ) {{
      const val = String(item[key]).trim();
      if (val !== "" && val !== "null" && val !== "undefined") {{
        return item[key];
      }}
    }}
  }}
  return "Sin usuario";
}}

document.addEventListener("DOMContentLoaded", () => {{
  initDateConstraints();
  updateDependentFiltersState();
  applyFiltersAndUpdateOptions();

  document.querySelectorAll(".navbtn").forEach(btn => {{
    btn.addEventListener("click", (e) => {{
      document.querySelectorAll(".navbtn").forEach(b => b.classList.remove("active"));
      const target = e.currentTarget;
      target.classList.add("active");
      
      currentExec = target.getAttribute("data-exec");
      document.getElementById("pageTitle").innerText = currentExec;
      
      currentPage = 1;
      updateDependentFiltersState();
      applyFiltersAndUpdateOptions();
    }});
  }});

  const filterBar = document.getElementById("filterBar");
  const toggleBtn = document.getElementById("toggleFiltersBtn");

  toggleBtn.addEventListener("click", () => {{
    const isCollapsed = filterBar.classList.toggle("collapsed");
    toggleBtn.innerText = isCollapsed ? "Mostrar filtros" : "Ocultar filtros";
    toggleBtn.setAttribute("aria-expanded", !isCollapsed);
  }});

  ["fDateFrom", "fDateTo"].forEach(id => {{
    const el = document.getElementById(id);
    if (el) {{
      el.addEventListener("change", () => {{ currentPage = 1; applyFiltersAndUpdateOptions(); }});
      el.addEventListener("input", () => {{ currentPage = 1; applyFiltersAndUpdateOptions(); }});
    }}
  }});

  const filterInputs = ["fDistrict", "fChannel", "fChain", "fCity", "fPuntoVenta", "fCategory", "fExhibType"];
  
  filterInputs.forEach(id => {{
    const input = document.getElementById(id);
    const dropdown = document.getElementById("drop_" + id);

    input.addEventListener("focus", () => {{
      if (!input.disabled) {{
        renderDropdownOptions(id, input.value);
        dropdown.classList.add("active");
      }}
    }});

    input.addEventListener("input", () => {{
      renderDropdownOptions(id, input.value);
      dropdown.classList.add("active");
      currentPage = 1;
      applyFiltersAndUpdateOptions(id);
    }});
  }});

  document.addEventListener("click", (e) => {{
    if (!e.target.closest(".searchable-select-wrapper")) {{
      document.querySelectorAll(".searchable-dropdown").forEach(d => d.classList.remove("active"));
    }}
  }});

  document.getElementById("clearBtn").addEventListener("click", () => {{
    document.getElementById("fDateFrom").value = "";
    document.getElementById("fDateTo").value = "";
    filterInputs.forEach(id => {{
      const el = document.getElementById(id);
      if (!el.disabled) el.value = "";
    }});
    currentPage = 1;
    applyFiltersAndUpdateOptions();
  }});

  document.getElementById("exportBtn").addEventListener("click", exportToCSV);

  document.getElementById("closeModalBtn").addEventListener("click", closeModal);
  document.getElementById("lightboxModal").addEventListener("click", (e) => {{
    if (e.target.id === "lightboxModal") closeModal();
  }});
}});

function updateDependentFiltersState() {{
  const categoryInput = document.getElementById("fCategory");
  const exhibInput = document.getElementById("fExhibType");

  const isSurtido = currentExec === "Surtido";
  const isExhibicion = currentExec === "Exhibición adicional" || currentExec === "Exhibición paga / convenio";

  if (isSurtido) {{
    categoryInput.disabled = false;
    categoryInput.placeholder = "Todas las categorías";
  }} else {{
    categoryInput.disabled = true;
    categoryInput.value = "";
    categoryInput.placeholder = "N/A para esta sección";
  }}

  if (isExhibicion) {{
    exhibInput.disabled = false;
    exhibInput.placeholder = "Todas las exhibiciones";
  }} else {{
    exhibInput.disabled = true;
    exhibInput.value = "";
    exhibInput.placeholder = "N/A para esta sección";
  }}
}}

function matchesFilters(item, excludeKey = null) {{
  const dateFrom = document.getElementById("fDateFrom").value;
  const dateTo = document.getElementById("fDateTo").value;
  const districtVal = document.getElementById("fDistrict").value.trim().toLowerCase();
  const channelVal = document.getElementById("fChannel").value.trim().toLowerCase();
  const chainVal = document.getElementById("fChain").value.trim().toLowerCase();
  const cityVal = document.getElementById("fCity").value.trim().toLowerCase();
  const pdvVal = document.getElementById("fPuntoVenta").value.trim().toLowerCase();
  const categoryVal = document.getElementById("fCategory").value.trim().toLowerCase();
  const exhibTypeVal = document.getElementById("fExhibType").value.trim().toLowerCase();

  const itemDate = item["Fecha"] || "";
  const matchesDateFrom = !dateFrom || itemDate >= dateFrom;
  const matchesDateTo = !dateTo || itemDate <= dateTo;

  const matchField = (itemVal, filterVal) => {{
    if (!filterVal) return true;
    return String(itemVal || "").toLowerCase().includes(filterVal);
  }};

  const matchesDistrict = excludeKey === "Distrito" || matchField(item["Distrito"], districtVal);
  const matchesChannel = excludeKey === "Canal" || matchField(item["Canal"], channelVal);
  const matchesChain = excludeKey === "Cadena" || matchField(item["Cadena"], chainVal);
  const matchesCity = excludeKey === "Ciudad" || matchField(item["Ciudad"], cityVal);
  const matchesPdv = excludeKey === "Punto de Venta" || matchField(item["Punto de Venta"], pdvVal);
  const matchesCategory = excludeKey === "Respuesta" || matchField(item["Respuesta"], categoryVal);
  const matchesExhibType = excludeKey === "RespuestaExhib" || matchField(item["Respuesta"], exhibTypeVal);

  const fotoAntes = item["Foto antes"] ? String(item["Foto antes"]).trim() : "";
  const fotoDespues = item["Foto despues"] ? String(item["Foto despues"]).trim() : "";
  const hasPhoto = fotoAntes !== "" || fotoDespues !== "";

  return matchesDateFrom && matchesDateTo && matchesDistrict && matchesChannel && matchesChain && matchesCity && matchesPdv && matchesCategory && matchesExhibType && hasPhoto;
}}

function getFilteredData() {{
  return rawData.filter(item => {{
    const matchesExec = String(item["Ejecucion"] || "").trim() === currentExec;
    return matchesExec && matchesFilters(item);
  }});
}}

function getAvailableValues(key, excludeKey) {{
  return [...new Set(
    rawData
      .filter(item => {{
        const matchesExec = String(item["Ejecucion"] || "").trim() === currentExec;
        return matchesExec && matchesFilters(item, excludeKey);
      }})
      .map(item => String(item[key] || "").trim())
      .filter(Boolean)
  )].sort();
}}

function renderDropdownOptions(id, searchTerm = "") {{
  const dropdown = document.getElementById("drop_" + id);
  dropdown.innerHTML = "";

  let key = "";
  let excludeKey = null;
  let defaultLabel = "";

  if (id === "fDistrict") {{ key = "Distrito"; excludeKey = "Distrito"; defaultLabel = "Todos los distritos"; }}
  else if (id === "fChannel") {{ key = "Canal"; excludeKey = "Canal"; defaultLabel = "Todos los canales"; }}
  else if (id === "fChain") {{ key = "Cadena"; excludeKey = "Cadena"; defaultLabel = "Todas las cadenas"; }}
  else if (id === "fCity") {{ key = "Ciudad"; excludeKey = "Ciudad"; defaultLabel = "Todas las ciudades"; }}
  else if (id === "fPuntoVenta") {{ key = "Punto de Venta"; excludeKey = "Punto de Venta"; defaultLabel = "Todos los puntos"; }}
  else if (id === "fCategory") {{ key = "Respuesta"; excludeKey = "Respuesta"; defaultLabel = "Todas las categorías"; }}
  else if (id === "fExhibType") {{ key = "Respuesta"; excludeKey = "RespuestaExhib"; defaultLabel = "Todas las exhibiciones"; }}

  const values = getAvailableValues(key, excludeKey);
  const term = searchTerm.toLowerCase().trim();
  const filteredVals = values.filter(v => v.toLowerCase().includes(term));

  const defaultDiv = document.createElement("div");
  defaultDiv.className = "searchable-option";
  defaultDiv.style.fontWeight = "bold";
  defaultDiv.style.color = "var(--muted)";
  defaultDiv.textContent = defaultLabel;
  defaultDiv.onclick = () => {{
    document.getElementById(id).value = "";
    dropdown.classList.remove("active");
    currentPage = 1;
    applyFiltersAndUpdateOptions();
  }};
  dropdown.appendChild(defaultDiv);

  if (filteredVals.length === 0) {{
    const noRes = document.createElement("div");
    noRes.className = "searchable-option";
    noRes.style.color = "#94a3b8";
    noRes.style.cursor = "default";
    noRes.textContent = "No se encontraron resultados";
    dropdown.appendChild(noRes);
  }} else {{
    filteredVals.forEach(val => {{
      const optDiv = document.createElement("div");
      optDiv.className = "searchable-option";
      optDiv.textContent = val;
      optDiv.onclick = () => {{
        document.getElementById(id).value = val;
        dropdown.classList.remove("active");
        currentPage = 1;
        applyFiltersAndUpdateOptions();
      }};
      dropdown.appendChild(optDiv);
    }});
  }}
}}

function applyFiltersAndUpdateOptions(changedId = null) {{
  if (currentExec !== "Surtido") {{
    document.getElementById("fCategory").value = "";
  }}
  if (currentExec !== "Exhibición adicional" && currentExec !== "Exhibición paga / convenio") {{
    document.getElementById("fExhibType").value = "";
  }}

  renderCards();
}}

function changePage(direction) {{
  const filtered = getFilteredData();
  const totalPages = Math.ceil(filtered.length / pageSize) || 1;
  
  currentPage += direction;
  if (currentPage < 1) currentPage = 1;
  if (currentPage > totalPages) currentPage = totalPages;
  
  renderCards();
  window.scrollTo({{ top: 0, behavior: 'smooth' }});
}}

function updateActiveFilterChips() {{
  const container = document.getElementById("activeFilters");
  container.innerHTML = '<span class="caption">Filtros activos:</span>';

  const addChip = (label, val) => {{
    if (val && val.trim() !== "") {{
      const chip = document.createElement("span");
      chip.className = "filter-chip";
      chip.textContent = `${{label}}: ${{val}}`;
      container.appendChild(chip);
    }}
  }};

  addChip("Desde", document.getElementById("fDateFrom").value);
  addChip("Hasta", document.getElementById("fDateTo").value);
  addChip("Distrito", document.getElementById("fDistrict").value);
  addChip("Canal", document.getElementById("fChannel").value);
  addChip("Cadena", document.getElementById("fChain").value);
  addChip("Ciudad", document.getElementById("fCity").value);
  addChip("Punto de venta", document.getElementById("fPuntoVenta").value);
  
  if (currentExec === "Surtido") {{
    addChip("Categoría", document.getElementById("fCategory").value);
  }}
  
  if (currentExec === "Exhibición adicional" || currentExec === "Exhibición paga / convenio") {{
    addChip("Tipo de exhibición", document.getElementById("fExhibType").value);
  }}
}}

function renderCards() {{
  const grid = document.getElementById("photoGrid");
  const paginationContainer = document.getElementById("paginationContainer");
  grid.innerHTML = "";

  const filtered = getFilteredData();
  updateActiveFilterChips();

  if (filtered.length === 0) {{
    grid.innerHTML = '<p style="grid-column: 1/-1; text-align:center; color: var(--muted); padding: 40px 0;">No se encontraron registros con fotos para la selección actual.</p>';
    paginationContainer.style.display = "none";
    return;
  }}

  const totalPages = Math.ceil(filtered.length / pageSize);
  if (currentPage > totalPages) currentPage = totalPages;

  const startIndex = (currentPage - 1) * pageSize;
  const endIndex = startIndex + pageSize;
  const paginatedItems = filtered.slice(startIndex, endIndex);

  paginatedItems.forEach((item) => {{
    const card = document.createElement("div");
    card.className = "card";
    card.onclick = () => openModal(item);

    const isSurtido = currentExec === "Surtido";
    const fotoAntes = item["Foto antes"] ? String(item["Foto antes"]).trim() : "";
    const fotoDespues = item["Foto despues"] ? String(item["Foto despues"]).trim() : "";

    let imageContent = "";
    if (isSurtido && fotoAntes && fotoDespues) {{
      imageContent = `
        <div class="card-img-wrapper dual">
          <span class="card-badge">Antes / Después</span>
          <img src="${{fotoAntes}}" alt="Antes" loading="lazy"/>
          <img src="${{fotoDespues}}" alt="Después" loading="lazy"/>
        </div>`;
    }} else {{
      const imgToShow = fotoDespues || fotoAntes;
      imageContent = `
        <div class="card-img-wrapper">
          <img src="${{imgToShow}}" alt="Foto Ejecución" loading="lazy"/>
        </div>`;
    }}

    const usuario = getUserName(item);
    const ciudad = item["Ciudad"] || "Sin ciudad";
    const fecha = item["Fecha"] || "N/A";
    const punto = item["Punto de Venta"] || "PDV Sin Nombre";
    
    const respuestaEjecucion = item["Respuesta"] || item["Cantidad"] || "";

    const isExhibicion = currentExec === "Exhibición adicional" || currentExec === "Exhibición paga / convenio";
    const tipoHtml = isExhibicion ? `<div class="card-info-item full-width"><span class="icon">📦</span> <span>Tipo: ${{respuestaEjecucion}}</span></div>` : "";
    const catHtml = isSurtido && respuestaEjecucion ? `<div class="card-info-item full-width"><span class="icon">🏷️</span> <span>Categoría: ${{respuestaEjecucion}}</span></div>` : "";

    card.innerHTML = `
      ${{imageContent}}
      <div class="card-body">
        <h3 class="card-title">${{punto}}</h3>
        <div class="card-info-grid">
          <div class="card-info-item"><span class="icon">👤</span> <span>${{usuario}}</span></div>
          <div class="card-info-item"><span class="icon">🌆</span> <span>${{ciudad}}</span></div>
          <div class="card-info-item"><span class="icon">📅</span> <span>${{fecha}}</span></div>
          ${{catHtml}}
          ${{tipoHtml}}
        </div>
      </div>
    `;

    grid.appendChild(card);
  }});

  paginationContainer.style.display = "flex";
  document.getElementById("paginationInfo").innerText = `Mostrando ${{startIndex + 1}} - ${{Math.min(endIndex, filtered.length)}} de ${{filtered.length}} registros (Página ${{currentPage}} de ${{totalPages}})`;
  document.getElementById("prevPageBtn").disabled = currentPage === 1;
  document.getElementById("nextPageBtn").disabled = currentPage === totalPages;
  document.getElementById("prevPageBtn").style.opacity = currentPage === 1 ? "0.5" : "1";
  document.getElementById("nextPageBtn").style.opacity = currentPage === totalPages ? "0.5" : "1";
}}

function openModal(item) {{
  const modal = document.getElementById("lightboxModal");
  const modalImg = document.getElementById("modalImg");
  const thumbsContainer = document.getElementById("modalThumbnails");

  document.getElementById("mPuntoVenta").innerText = item["Punto de Venta"] || "N/A";
  document.getElementById("mNombreUsuario").innerText = getUserName(item);
  document.getElementById("mCiudad").innerText = item["Ciudad"] || "N/A";
  document.getElementById("mFecha").innerText = item["Fecha"] || "N/A";
  
  const respuestaEjecucion = item["Respuesta"] || item["Cantidad"] || "";
  const isExhibicion = currentExec === "Exhibición adicional" || currentExec === "Exhibición paga / convenio";
  const isSurtido = currentExec === "Surtido";
  const mTipoContainer = document.getElementById("mTipoContainer");
  const mTipoLabel = document.getElementById("mTipoLabel");

  if (isExhibicion && respuestaEjecucion) {{
    mTipoContainer.style.display = "flex";
    mTipoLabel.innerText = "Tipo";
    document.getElementById("mTipo").innerText = respuestaEjecucion;
  }} else if (isSurtido && respuestaEjecucion) {{
    mTipoContainer.style.display = "flex";
    mTipoLabel.innerText = "Categoría";
    document.getElementById("mTipo").innerText = respuestaEjecucion;
  }} else {{
    mTipoContainer.style.display = "none";
  }}

  thumbsContainer.innerHTML = "";

  const fotoAntes = item["Foto antes"] ? String(item["Foto antes"]).trim() : "";
  const fotoDespues = item["Foto despues"] ? String(item["Foto despues"]).trim() : "";

  let photos = [];
  if (isSurtido) {{
    if (fotoAntes) photos.push({{ type: "Antes", url: fotoAntes }});
    if (fotoDespues) photos.push({{ type: "Después", url: fotoDespues }});
  }} else {{
    if (fotoDespues) photos.push({{ type: "Después", url: fotoDespues }});
    else if (fotoAntes) photos.push({{ type: "Antes", url: fotoAntes }});
  }}

  document.getElementById("mPhotosTitle").innerText = `FOTOS DEL REGISTRO (${{photos.length}})`;

  if (photos.length > 0) {{
    modalImg.src = photos[0].url;

    photos.forEach((photo, idx) => {{
      const wrapper = document.createElement("div");
      wrapper.className = "thumb-item";

      const btn = document.createElement("button");
      btn.className = `thumb-btn ${{idx === 0 ? "active" : ""}}`;
      btn.innerHTML = `<img src="${{photo.url}}" alt="${{photo.type}}"/>`;
      btn.onclick = () => {{
        document.querySelectorAll(".thumb-btn").forEach(b => b.classList.remove("active"));
        btn.classList.add("active");
        modalImg.src = photo.url;
      }};

      const label = document.createElement("span");
      label.className = "thumb-label";
      label.textContent = photo.type;

      wrapper.appendChild(btn);
      wrapper.appendChild(label);
      thumbsContainer.appendChild(wrapper);
    }});
  }}

  modal.classList.add("active");
}}

function closeModal() {{
  document.getElementById("lightboxModal").classList.remove("active");
}}

function exportToCSV() {{
  const data = getFilteredData();
  if (!data || !data.length) {{
    alert("No hay datos para exportar.");
    return;
  }}

  const headers = Object.keys(data[0]);
  const csvRows = [];
  csvRows.push(headers.join(";"));

  data.forEach(row => {{
    const values = headers.map(header => {{
      const val = row[header] === null || row[header] === undefined ? "" : String(row[header]);
      return `"${{val.replace(/"/g, '""')}}"`;
    }});
    csvRows.push(values.join(";"));
  }});

  const csvContent = "\\uFEFF" + csvRows.join("\\n");
  const blob = new Blob([csvContent], {{ type: "text/csv;charset=utf-8;" }});
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.setAttribute("href", url);
  link.setAttribute("download", `Ejecucion_PDV_${{currentExec}}_${{new Date().toISOString().slice(0,10)}}.csv`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}}
</script>
</body>
</html>
"""

# Guardar el archivo index.html
with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)

print(f"¡Archivo index.html generado con éxito sin animación! Fecha de corte detectada: {fecha_corte_str}")
input("\nPresiona Enter para cerrar la consola...")