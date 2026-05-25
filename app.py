import streamlit as st
import pandas as pd
import random
from io import BytesIO
import datetime

# ── CONFIGURACIÓN ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Sistema MRP + VRP — NETO DURANGAR S.A.S.",
    page_icon="🚚", layout="wide"
)

st.markdown("""
<style>
.titulo { background:linear-gradient(90deg,#1F3864,#2E75B6);
          color:white;padding:18px;border-radius:10px;text-align:center;margin-bottom:16px; }
.tab-desc { color:#555; font-size:14px; margin-bottom:12px; }
.kpi-card { background:#f0f4f8; border-left:4px solid #2E75B6;
            padding:12px;border-radius:6px;margin:6px 0; }
.ruta-ok   { background:#E2EFDA;border-left:4px solid #375623;
             padding:10px;border-radius:6px;margin:6px 0; }
.ruta-warn { background:#FFF2CC;border-left:4px solid #F4A300;
             padding:10px;border-radius:6px;margin:6px 0; }
.ruta-bad  { background:#FCE4D6;border-left:4px solid #C55A11;
             padding:10px;border-radius:6px;margin:6px 0; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="titulo">
  <h2 style="margin:0">🍽️ Sistema Integrado MRP + VRP — NETO DURANGAR S.A.S.</h2>
  <p style="margin:4px 0 0;opacity:.9">Generador de Menús · Requerimientos de Materiales · Rutas Clarke & Wright</p>
</div>""", unsafe_allow_html=True)

# ── DATOS FIJOS VRP (tomados del Excel ClarkeWright_VRP_Durangar.xlsx) ────────
CAMPOS = ['Base H&P','Carrizales','Yenac','Careto','Corcel','Arrendajo']

DIST_BODEGA = {
    'Base H&P':36,'Carrizales':90,'Yenac':119,
    'Careto':158,'Corcel':178,'Arrendajo':210
}
T_VIAJE_BODEGA = {
    'Base H&P':60,'Carrizales':159,'Yenac':190,
    'Careto':235,'Corcel':210,'Arrendajo':330
}
T_DESCARGA = 65  # min por campo

DIST_INTER = {
    ('Base H&P','Carrizales'):85, ('Base H&P','Yenac'):123,
    ('Base H&P','Careto'):174,    ('Base H&P','Corcel'):200,
    ('Base H&P','Arrendajo'):241,
    ('Carrizales','Yenac'):53,    ('Carrizales','Careto'):103,
    ('Carrizales','Corcel'):129,  ('Carrizales','Arrendajo'):171,
    ('Yenac','Careto'):66,        ('Yenac','Corcel'):92,
    ('Yenac','Arrendajo'):133,
    ('Careto','Corcel'):41,       ('Careto','Arrendajo'):83,
    ('Corcel','Arrendajo'):57,
}
def dist_inter(a, b):
    if a == b: return 0
    return DIST_INTER.get((a,b), DIST_INTER.get((b,a), 0))

VEHICULOS = pd.DataFrame([
    {'placa':'TSN320','tipo':'Camión',    'costo_km':1387,'cap_kg':2000,'refrig':True},
    {'placa':'WEQ',   'tipo':'Camión',    'costo_km':675, 'cap_kg':8000,'refrig':True},
    {'placa':'JTX761','tipo':'Camioneta', 'costo_km':221, 'cap_kg':700, 'refrig':False},
    {'placa':'SXD',   'tipo':'Camioneta', 'costo_km':204, 'cap_kg':700, 'refrig':False},
    {'placa':'SZO',   'tipo':'Camioneta', 'costo_km':253, 'cap_kg':700, 'refrig':False},
])

COSTO_KM_CAMION    = 767   # $/km promedio camión diesel
COSTO_KM_CAMIONETA = 218   # $/km promedio camioneta diesel

# ── CARGAR RECETAS ────────────────────────────────────────────────────────────
@st.cache_data
def cargar_recetas():
    df_des = pd.read_excel("Desayuno_Final_02.xlsx", header=0)
    df_des.columns = ['_','COMIDA','DESCRIPCION','PREPARACION',
                      'INGREDIENTE','g_persona','num_personas','total_dia','total_mes']
    df_des = df_des.iloc[1:].dropna(subset=['PREPARACION','INGREDIENTE'])
    df_des['g_persona'] = pd.to_numeric(df_des['g_persona'], errors='coerce')

    df_alm = pd.read_excel("Alm-Cena_02.xlsx", header=0)
    df_alm.columns = ['_','DESCRIPCION','PREPARACION','INGREDIENTE',
                      'g_persona','num_personas','total_dia','total_mes']
    df_alm = df_alm.iloc[1:].dropna(subset=['PREPARACION','INGREDIENTE'])
    df_alm['g_persona'] = pd.to_numeric(df_alm['g_persona'], errors='coerce')
    return df_des, df_alm

df_des, df_alm = cargar_recetas()

# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Parámetros generales")
    num_personas  = st.number_input("Personas/servicio", 100, 2000, 690, 10)
    num_semanas   = st.slider("Semanas del ciclo", 1, 4, 4)
    margen_pct    = st.slider("Margen seguridad MP (%)", 0, 20, 10)
    ventana_min   = st.selectbox("Ventana horaria (min)", [300, 540], index=1,
                                  format_func=lambda x: f"{'5am–10am' if x==300 else '5am–2pm'} ({x} min)")
    st.markdown("---")
    st.markdown("### 🚚 Flota disponible")
    st.dataframe(VEHICULOS[['placa','tipo','costo_km','cap_kg']],
                 hide_index=True, use_container_width=True)

# ── TABS ──────────────────────────────────────────────────────────────────────
t1, t2, t3 = st.tabs([
    "🍽️ 1. Generador de Menús",
    "📦 2. MRP — Requerimientos",
    "🚚 3. VRP — Rutas Clarke & Wright"
])

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 1 — GENERADOR DE MENÚS
# ═══════════════════════════════════════════════════════════════════════════════
dias = ['Lunes','Martes','Miércoles','Jueves','Viernes','Sábado','Domingo']

def elegir(df, desc, excluir=None):
    ops = df[df['DESCRIPCION']==desc]['PREPARACION'].unique().tolist()
    if excluir: ops = [x for x in ops if x not in excluir] or ops
    return random.choice(ops) if ops else "N/D"

def generar_menu(semanas, personas, rotar):
    menu = {}
    j_des, j_alm = [], []
    for s in range(1, semanas+1):
        menu[f'Semana {s}'] = {}
        for dia in dias:
            jd = elegir(df_des,'Jugos', j_des if rotar else None)
            if rotar: j_des.append(jd); j_des[:] = j_des[-10:]
            ja = elegir(df_alm,'Jugos', j_alm if rotar else None)
            if rotar: j_alm.append(ja); j_alm[:] = j_alm[-10:]

            prot1_alm = (elegir(df_alm,'Especial - 1x Semana (Pescados y Mariscos)')
                         if dia=='Viernes'
                         else elegir(df_alm,'Proteico 1 - Carne Roja'))

            menu[f'Semana {s}'][dia] = {
                'Desayuno': {
                    'Jugo':           jd,
                    'Bebida caliente':elegir(df_des,'Bebida Caliente '),
                    'Lácteo':         elegir(df_des,'Lácteos'),
                    'Fruta':          elegir(df_des,'Fruta'),
                    'Cereal':         elegir(df_des,'Cereales'),
                    'Queso':          elegir(df_des,'Queso'),
                    'Huevo':          elegir(df_des,'Huevos '),
                    'Proteína 1':     elegir(df_des,'Proteina 1'),
                    'Proteína 2':     elegir(df_des,'Proteina 2'),
                    'Caldo':          elegir(df_des,'Caldo'),
                    'Arroz':          elegir(df_des,'Arroz Cocido'),
                    'Pan':            elegir(df_des,'Pan Varios'),
                },
                'Almuerzo': {
                    'Jugo':           ja,
                    'Fruta de mano':  elegir(df_alm,'Frutas de Mano'),
                    'Sopa/Crema':     elegir(df_alm,'Sopa, Crema o Consomé'),
                    'Proteína 1':     prot1_alm,
                    'Proteína 2':     elegir(df_alm,'Proteico 2 - Carne Blanca'),
                    'Verduras':       elegir(df_alm,'Verduras Cocidas'),
                    'Arroz':          elegir(df_alm,'Arroz'),
                    'Energético':     elegir(df_alm,'Energético'),
                    'Ensalada':       elegir(df_alm,'Barra de Ensalada'),
                    'Leguminosa':     elegir(df_alm,'Leguminosa'),
                    'Postre':         elegir(df_alm,'Postre'),
                },
                'Cena': {
                    'Jugo':           elegir(df_alm,'Jugos',[ja]),
                    'Fruta de mano':  elegir(df_alm,'Frutas de Mano'),
                    'Sopa/Crema':     elegir(df_alm,'Sopa, Crema o Consomé'),
                    'Proteína 1':     elegir(df_alm,'Proteico 1 - Carne Roja'),
                    'Proteína 2':     elegir(df_alm,'Proteico 2 - Carne Blanca'),
                    'Verduras':       elegir(df_alm,'Verduras Cocidas'),
                    'Arroz':          elegir(df_alm,'Arroz'),
                    'Energético':     elegir(df_alm,'Energético'),
                    'Ensalada':       elegir(df_alm,'Barra de Ensalada'),
                    'Leguminosa':     elegir(df_alm,'Leguminosa'),
                    'Postre':         elegir(df_alm,'Postre'),
                }
            }
    return menu

with t1:
    st.markdown('<p class="tab-desc">Genera un menú aleatorio completo respetando la Minuta Contractual de Frontera Energy.</p>', unsafe_allow_html=True)
    rotar = st.checkbox("Rotar jugos (sin repetir consecutivos)", value=True)

    col1, col2, col3 = st.columns([1,2,1])
    with col2:
        if st.button("🎲 GENERAR MENÚ ALEATORIO", type="primary", use_container_width=True):
            st.session_state['menu'] = generar_menu(num_semanas, num_personas, rotar)
            st.session_state['params_menu'] = {'personas':num_personas,'margen':margen_pct,'semanas':num_semanas}

    if 'menu' in st.session_state:
        menu = st.session_state['menu']
        c1,c2,c3,c4 = st.columns(4)
        c1.metric("Semanas",num_semanas)
        c2.metric("Personas/servicio",num_personas)
        c3.metric("Días generados",num_semanas*7)
        c4.metric("Servicios totales",num_semanas*7*3)

        st.markdown("---")
        sem_sel = st.selectbox("Ver semana:", list(menu.keys()))
        for dia, servicios in menu[sem_sel].items():
            with st.expander(f"📅 {dia}", expanded=(dia=='Lunes')):
                cd, ca, cc = st.columns(3)
                with cd:
                    st.markdown("**🌅 Desayuno**")
                    for k,v in servicios['Desayuno'].items():
                        st.markdown(f"- **{k}:** {v}")
                with ca:
                    st.markdown("**☀️ Almuerzo**")
                    for k,v in servicios['Almuerzo'].items():
                        st.markdown(f"- **{k}:** {v}")
                with cc:
                    st.markdown("**🌙 Cena**")
                    for k,v in servicios['Cena'].items():
                        st.markdown(f"- **{k}:** {v}")

        # Exportar Excel
        def exportar_menu_excel(menu, personas, margen):
            out = BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as w:
                for sem, dias_m in menu.items():
                    filas = []
                    for dia, servicios in dias_m.items():
                        for serv, preps in servicios.items():
                            for comp, prep in preps.items():
                                df_src = df_des if serv=='Desayuno' else df_alm
                                ing = df_src[df_src['PREPARACION']==prep][['INGREDIENTE','g_persona']].drop_duplicates()
                                if len(ing)==0:
                                    filas.append({'DIA':dia,'SERVICIO':serv,'COMPONENTE':comp,
                                                  'PREPARACION':prep,'INGREDIENTE':'','g/persona':0,
                                                  '# PERSONAS':personas,'TOTAL Kg':0,
                                                  f'TOTAL Kg +{margen}%':0})
                                else:
                                    for _, r in ing.iterrows():
                                        g = float(r['g_persona']) if pd.notna(r['g_persona']) else 0
                                        tb = round(g*personas/1000,3)
                                        filas.append({'DIA':dia,'SERVICIO':serv,'COMPONENTE':comp,
                                                      'PREPARACION':prep,'INGREDIENTE':r['INGREDIENTE'],
                                                      'g/persona':g,'# PERSONAS':personas,
                                                      'TOTAL Kg':tb,
                                                      f'TOTAL Kg +{margen}%':round(tb*(1+margen/100),3)})
                    pd.DataFrame(filas).to_excel(w, sheet_name=sem.replace(' ','_')[:31], index=False)
            out.seek(0); return out

        st.markdown("---")
        fecha = datetime.datetime.now().strftime("%Y%m%d_%H%M")
        col_a, col_b = st.columns(2)
        with col_a:
            st.download_button("📊 Descargar Excel del menú",
                data=exportar_menu_excel(menu, num_personas, margen_pct),
                file_name=f"Menu_DURANGAR_{fecha}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)
        with col_b:
            if st.button("🔄 Regenerar menú", use_container_width=True):
                st.session_state['menu'] = generar_menu(num_semanas, num_personas, rotar)
                st.rerun()
    else:
        st.info("Configura los parámetros en el panel izquierdo y haz clic en **GENERAR MENÚ ALEATORIO**.")

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 2 — MRP
# ═══════════════════════════════════════════════════════════════════════════════
with t2:
    st.markdown('<p class="tab-desc">Calcula el requerimiento total de materia prima a partir del menú generado.</p>', unsafe_allow_html=True)

    if 'menu' not in st.session_state:
        st.warning("Primero genera un menú en la pestaña 1.")
    else:
        menu = st.session_state['menu']
        personas = num_personas
        margen = margen_pct / 100

        # Consolidar ingredientes de todo el menú
        filas_mrp = []
        for sem, dias_m in menu.items():
            for dia, servicios in dias_m.items():
                for serv, preps in servicios.items():
                    for comp, prep in preps.items():
                        df_src = df_des if serv=='Desayuno' else df_alm
                        ing = df_src[df_src['PREPARACION']==prep][['INGREDIENTE','g_persona']].drop_duplicates()
                        for _, r in ing.iterrows():
                            g = float(r['g_persona']) if pd.notna(r['g_persona']) else 0
                            filas_mrp.append({
                                'semana': sem, 'dia': dia, 'servicio': serv,
                                'ingrediente': r['INGREDIENTE'],
                                'g_persona': g,
                                'total_kg': round(g * personas / 1000, 3)
                            })

        df_mrp_det = pd.DataFrame(filas_mrp)

        # Resumen por ingrediente
        mrp_res = df_mrp_det.groupby('ingrediente')['total_kg'].sum().reset_index()
        mrp_res.columns = ['Ingrediente','Total base (kg)']
        mrp_res['Margen 10% (kg)'] = (mrp_res['Total base (kg)'] * margen).round(3)
        mrp_res['Pedido total (kg)'] = (mrp_res['Total base (kg)'] * (1+margen)).round(3)
        mrp_res = mrp_res.sort_values('Pedido total (kg)', ascending=False).reset_index(drop=True)

        # KPIs
        c1,c2,c3,c4 = st.columns(4)
        c1.metric("Ingredientes únicos", len(mrp_res))
        c2.metric("Base total (kg)", f"{mrp_res['Total base (kg)'].sum():,.1f}")
        c3.metric(f"Margen {margen_pct}% (kg)", f"{mrp_res['Margen 10% (kg)'].sum():,.1f}")
        c4.metric("PEDIDO TOTAL (kg)", f"{mrp_res['Pedido total (kg)'].sum():,.1f}")

        st.markdown("---")
        st.markdown("#### Requerimiento por ingrediente")
        st.dataframe(mrp_res, use_container_width=True, hide_index=True)

        # Guardar para VRP
        st.session_state['mrp'] = mrp_res

        # Exportar MRP
        out_mrp = BytesIO()
        with pd.ExcelWriter(out_mrp, engine='openpyxl') as w:
            mrp_res.to_excel(w, sheet_name='MRP_Requerimientos', index=False)
            df_mrp_det.to_excel(w, sheet_name='Detalle_por_servicio', index=False)
        out_mrp.seek(0)
        fecha = datetime.datetime.now().strftime("%Y%m%d_%H%M")
        st.download_button("📊 Descargar MRP Excel",
            data=out_mrp,
            file_name=f"MRP_DURANGAR_{fecha}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 3 — VRP CLARKE & WRIGHT
# ═══════════════════════════════════════════════════════════════════════════════
with t3:
    st.markdown('<p class="tab-desc">Optimización de rutas de distribución a campos petroleros usando el método de Ahorros de Clarke & Wright.</p>', unsafe_allow_html=True)

    # Demanda por campo (editable)
    st.markdown("#### Demanda quincenal por campo (unidades)")
    demanda_def = {'Base H&P':378,'Carrizales':5970,'Yenac':4290,
                   'Careto':2500,'Corcel':11371,'Arrendajo':3589}

    col_dem = st.columns(3)
    demanda = {}
    campos_list = list(demanda_def.keys())
    for idx, campo in enumerate(campos_list):
        with col_dem[idx % 3]:
            demanda[campo] = st.number_input(campo, 0, 50000,
                                              demanda_def[campo], 100, key=f"dem_{campo}")

    st.markdown("---")

    if st.button("🚚 EJECUTAR VRP — CLARKE & WRIGHT", type="primary", use_container_width=True):

        # ── ALGORITMO CLARKE & WRIGHT ──────────────────────────────────────────
        # 1. Calcular ahorros S(i,j) = c(0,i) + c(0,j) - c(i,j)
        ahorros = []
        for i, ci in enumerate(campos_list):
            for j, cj in enumerate(campos_list):
                if j <= i: continue
                c0i = DIST_BODEGA[ci] * 2 * COSTO_KM_CAMION
                c0j = DIST_BODEGA[cj] * 2 * COSTO_KM_CAMION
                cij = dist_inter(ci, cj) * COSTO_KM_CAMION
                ahorro = c0i + c0j - cij
                # Tiempo de ruta combinada
                t_ruta = T_VIAJE_BODEGA[ci] + T_DESCARGA + dist_inter(ci,cj)*1.45 + T_VIAJE_BODEGA[cj] + T_DESCARGA
                factible_300 = t_ruta <= 300
                factible_540 = t_ruta <= 540
                ahorros.append({
                    'Campo i': ci, 'Campo j': cj,
                    'c(0,i) $': c0i, 'c(0,j) $': c0j,
                    'Dist. i→j (km)': dist_inter(ci,cj),
                    'c(i,j) $': cij,
                    'Ahorro S(i,j) $': ahorro,
                    'T. ruta (min)': round(t_ruta),
                    'Orden': f"Bodega→{ci}→{cj}→Bodega",
                    '✓ 300 min': '✓' if factible_300 else '✗',
                    '✓ 540 min': '✓' if factible_540 else '✗',
                })

        df_ahorros = pd.DataFrame(ahorros).sort_values('Ahorro S(i,j) $', ascending=False).reset_index(drop=True)
        df_ahorros.index += 1
        df_ahorros.index.name = 'Rank'

        # 2. Construir rutas (greedy)
        rutas_ind = []  # rutas individuales base
        for campo in campos_list:
            t_ind = T_VIAJE_BODEGA[campo] * 2 + T_DESCARGA
            km_ind = DIST_BODEGA[campo] * 2
            # Seleccionar vehículo
            if campo in ['Corcel']:
                veh = 'WEQ (Camión 8t refrig.)'
                ckm = 675
            elif campo in ['Carrizales','Yenac','Careto']:
                veh = 'WEQ (Camión refrig.)'
                ckm = 675
            elif campo == 'Arrendajo':
                veh = 'WEQ (Camión dedicado)'
                ckm = 675
            else:
                veh = 'SXD/JTX761 (Camioneta)'
                ckm = 218

            costo = km_ind * ckm
            rutas_ind.append({
                'Campo': campo,
                'Ruta': f"Bodega → {campo} → Bodega",
                'Km': km_ind, 'T. ruta (min)': t_ind,
                'Costo ($)': costo,
                'Vehículo': veh,
                '✓ 300 min': '✓' if t_ind<=300 else '✗',
                '✓ 540 min': '✓' if t_ind<=540 else '✗',
            })
        df_ind = pd.DataFrame(rutas_ind)

        # 3. Rutas propuestas (basadas en el Excel original)
        rutas_prop = [
            {'Ruta':'Ruta 1','Secuencia':'Bodega→Base H&P→Carrizales→Bodega',
             'Vehículo':'SXD (Camioneta diesel)','Km':211,'T. ruta (min)':472,
             'Costo ($)':211*204,'Demanda (und)':demanda['Base H&P']+demanda['Carrizales'],
             'Ahorro vs ind. ($)':31447,
             'Estado':'⚠ Ventana ampliada (540 min)'},
            {'Ruta':'Ruta 2','Secuencia':'Bodega→Yenac→Bodega',
             'Vehículo':'WEQ (Camión refrig.)','Km':238,'T. ruta (min)':445,
             'Costo ($)':238*675,'Demanda (und)':demanda['Yenac'],
             'Ahorro vs ind. ($)':0,
             'Estado':'✓ Factible 300 min'},
            {'Ruta':'Ruta 3','Secuencia':'Bodega→Careto→Bodega',
             'Vehículo':'WEQ (Camión refrig.)','Km':316,'T. ruta (min)':535,
             'Costo ($)':316*675,'Demanda (und)':demanda['Careto'],
             'Ahorro vs ind. ($)':0,
             'Estado':'⚠ Sin margen (exacto 540 min)'},
            {'Ruta':'Ruta 4','Secuencia':'Bodega→Corcel→Bodega',
             'Vehículo':'WEQ (Camión 8t refrig.)','Km':356,'T. ruta (min)':485,
             'Costo ($)':356*675,'Demanda (und)':demanda['Corcel'],
             'Ahorro vs ind. ($)':0,
             'Estado':'✓ Factible 300 min'},
            {'Ruta':'Ruta 5','Secuencia':'Bodega→Arrendajo→Bodega',
             'Vehículo':'WEQ (Camión dedicado)','Km':420,'T. ruta (min)':725,
             'Costo ($)':420*675,'Demanda (und)':demanda['Arrendajo'],
             'Ahorro vs ind. ($)':0,
             'Estado':'✗ Salida 3:25 am o ventana especial'},
        ]
        df_prop = pd.DataFrame(rutas_prop)

        # ── GUARDAR RESULTADOS ─────────────────────────────────────────────────
        st.session_state['vrp'] = {
            'ahorros': df_ahorros,
            'rutas_ind': df_ind,
            'rutas_prop': df_prop
        }

    # ── MOSTRAR RESULTADOS VRP ─────────────────────────────────────────────────
    if 'vrp' in st.session_state:
        vrp = st.session_state['vrp']
        df_ahorros = vrp['ahorros']
        df_prop    = vrp['rutas_prop']
        df_ind     = vrp['rutas_ind']

        # KPIs comparativo
        st.markdown("#### Comparativo: situación actual vs modelo VRP")
        c1,c2,c3,c4 = st.columns(4)
        km_act  = df_ind['Km'].sum()
        km_vrp  = df_prop['Km'].sum()
        costo_act = df_ind['Costo ($)'].sum()
        costo_vrp = df_prop['Costo ($)'].sum()
        c1.metric("Km actuales (quincenal)", f"{km_act:,}", f"{km_vrp-km_act:+,} km")
        c2.metric("Km modelo VRP",           f"{km_vrp:,}")
        c3.metric("Costo actual ($)",         f"${costo_act:,.0f}", f"${costo_vrp-costo_act:+,.0f}")
        c4.metric("Ahorro anual estimado",    f"${(costo_act-costo_vrp)*24:,.0f}")

        st.markdown("---")

        # Tabla de ahorros Clarke & Wright
        st.markdown("#### Tabla de ahorros S(i,j) — Clarke & Wright")
        st.markdown("**Fórmula:** S(i,j) = c(0,i) + c(0,j) − c(i,j)  |  Mayor ahorro = mayor prioridad de consolidar")
        factibles = df_ahorros[df_ahorros[f'✓ {ventana_min} min']=='✓']
        st.info(f"Pares factibles con ventana {ventana_min} min: **{len(factibles)}** par(es)")
        st.dataframe(df_ahorros, use_container_width=True)

        st.markdown("---")
        st.markdown("#### Rutas propuestas")
        for _, r in df_prop.iterrows():
            estado = r['Estado']
            cls = 'ruta-ok' if '✓' in estado else ('ruta-warn' if '⚠' in estado else 'ruta-bad')
            st.markdown(f"""
            <div class="{cls}">
              <strong>{r['Ruta']}</strong> — {r['Secuencia']}<br>
              🚛 {r['Vehículo']} &nbsp;|&nbsp; 📏 {r['Km']} km &nbsp;|&nbsp;
              ⏱ {r['T. ruta (min)']} min &nbsp;|&nbsp; 💰 ${r['Costo ($)']:,.0f}<br>
              📦 Demanda: {r['Demanda (und)']:,} und &nbsp;|&nbsp; {estado}
            </div>""", unsafe_allow_html=True)

        # Programación quincenal
        st.markdown("---")
        st.markdown("#### Programación quincenal de despachos")
        prog = pd.DataFrame([
            {'Ruta':'Ruta 1','Día':'Día 1 (Lunes)','Salida':'5:00 am',
             'Secuencia':'Bodega→Base H&P→Carrizales→Bodega',
             'Regreso':'12:33 pm','Vehículo':'SXD (Camioneta)','Estado':'⚠ Ventana ampliada'},
            {'Ruta':'Ruta 2','Día':'Día 1 (Lunes)','Salida':'5:00 am',
             'Secuencia':'Bodega→Yenac→Bodega',
             'Regreso':'~9:15 am','Vehículo':'WEQ (Camión refrig.)','Estado':'✓ OK'},
            {'Ruta':'Ruta 3','Día':'Día 2 (Martes)','Salida':'5:00 am',
             'Secuencia':'Bodega→Careto→Bodega',
             'Regreso':'10:00 am exacto','Vehículo':'WEQ (Camión refrig.)','Estado':'⚠ Sin margen'},
            {'Ruta':'Ruta 4','Día':'Día 2 (Martes)','Salida':'5:00 am',
             'Secuencia':'Bodega→Corcel→Bodega',
             'Regreso':'~9:35 am','Vehículo':'WEQ (Camión 8t refrig.)','Estado':'✓ OK'},
            {'Ruta':'Ruta 5','Día':'Día 3 (Miércoles)','Salida':'3:25 am*',
             'Secuencia':'Bodega→Arrendajo→Bodega',
             'Regreso':'~4:00 pm','Vehículo':'WEQ (Camión dedicado)','Estado':'✗ Salida especial'},
        ])
        st.dataframe(prog, use_container_width=True, hide_index=True)
        st.caption("* Ruta 5 (Arrendajo): opciones — (a) salida 3:25 am con hora extra, (b) pernocta conductor, (c) negociar ventana recepción 8am–1pm en campo Arrendajo.")

        # Exportar VRP
        out_vrp = BytesIO()
        with pd.ExcelWriter(out_vrp, engine='openpyxl') as w:
            df_ahorros.to_excel(w, sheet_name='Ahorros_CW', index=True)
            df_prop.to_excel(w, sheet_name='Rutas_Propuestas', index=False)
            df_ind.to_excel(w, sheet_name='Rutas_Individuales', index=False)
            prog.to_excel(w, sheet_name='Programacion_Quincenal', index=False)
        out_vrp.seek(0)
        fecha = datetime.datetime.now().strftime("%Y%m%d_%H%M")
        st.download_button("📊 Descargar VRP Excel completo",
            data=out_vrp,
            file_name=f"VRP_ClarkeWright_DURANGAR_{fecha}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
    else:
        st.info("Ajusta la demanda por campo y haz clic en **EJECUTAR VRP**.")

# ── FOOTER ────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    f"<p style='text-align:center;color:#888;font-size:12px'>"
    f"NETO DURANGAR S.A.S. | Sistema MRP + VRP — Tesis de Grado | "
    f"{datetime.datetime.now().strftime('%d/%m/%Y %H:%M')}</p>",
    unsafe_allow_html=True)
