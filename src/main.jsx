import React, { useState } from "react";
import { createRoot } from "react-dom/client";
import * as XLSX from "xlsx";
import { FileDown, Database, FileSpreadsheet, RotateCcw } from "lucide-react";
import "./styles.css";
import logo from "./logo-ficohsa.png";

const initial = {
  nombre: "Porcentaje de Cartera en Mora Mayor a 30 Días (M30)",
  codigo: "KRI-CRE-001",
  area: "Riesgo de Crédito",
  pais: "Honduras",
  tipo: "KRI",
  categoria: "Apetito de Riesgo / Regulatorio",
  riesgo: "Riesgo de Crédito",
  fechaCreacion: "Junio 2026",
  fechaActualizacion: "Junio 2026",
  version: "1.0",
  estado: "Activo",
  marco: "Normativa CNBS y lineamientos de Basilea aplicables.",
  alcance: "Honduras – Cartera de crédito comercial. Aplicable a Banco Ficohsa Honduras.",
  objetivo: "Monitorear el porcentaje de cartera con mora superior a 30 días respecto al saldo total de la cartera de crédito, permitiendo identificar oportunamente deterioros en la calidad de la cartera.",
  definicion: "Indicador que mide la proporción del saldo de cartera con atraso superior a 30 días respecto al saldo total de la cartera vigente.",
  interpretacion: "Incrementos en el indicador reflejan deterioro en la calidad de la cartera y una mayor exposición al riesgo de crédito.",
  unidad: "Porcentaje (%)",
  sentido: "Menor es mejor",
  formula: "M30 = (Saldo de cartera con mora > 30 días / Saldo total de cartera) × 100",
  reglas: "Considerar operaciones activas. Excluir operaciones castigadas. Utilizar saldos al cierre de cada período.",
  supuestos: "La información utilizada proviene de fuentes oficiales conciliadas por las áreas responsables.",
  variables: "Saldo Mora >30 | Saldo de créditos con atraso superior a 30 días | Core Bancario\nSaldo Total Cartera | Saldo total de créditos vigentes | Core Bancario",
  sistema: "Core Bancario",
  tabla: "Vista Consolidada de Cartera",
  responsableDato: "Gerencia de Administración de Crédito",
  freqAct: "Diaria",
  freqCalc: "Mensual",
  freqMon: "Mensual",
  freqRep: "Mensual",
  instancias: "Comité de Riesgos",
  apetito: "≤ 4.0%",
  tolerancia: "> 4.0% y ≤ 5.0%",
  capacidad: "> 5.0%",
  metodoUmbrales: "Comportamiento histórico, apetito de riesgo aprobado, tendencias observadas y juicio experto.",
  accionTol: "Analizar causas del deterioro y definir acciones preventivas.",
  accionCap: "Presentar plan de acción formal e informar a los comités correspondientes.",
  duenoIndicador: "Gerencia de Riesgo de Crédito",
  duenoDato: "Gerencia de Administración de Crédito",
  tecnico: "Analista de Riesgo de Crédito",
  revisor: "Gerencia de Gestión Integral de Riesgos",
  aprobador: "Vicepresidencia de Riesgos",
  validacionConceptual: "Conforme",
  validacionTecnica: "Conforme",
  validacionFuncional: "Conforme",
  limitaciones: "El indicador depende de la calidad y oportunidad de la información registrada en los sistemas fuente. Cambios regulatorios o metodológicos pueden afectar la comparabilidad histórica.",
  cambio: "Emisión inicial del indicador",
  elaboro: "Analista de Riesgo de Crédito",
  reviso: "Gerente de Riesgo de Crédito",
  aprobo: "Vicepresidente de Riesgos",
};

function Field({ label, name, value, setData, type = "input", options = [] }) {
  const update = (e) => setData((d) => ({ ...d, [name]: e.target.value }));

  return (
    <div className={type === "textarea" ? "full" : ""}>
      <label>{label}</label>
      {type === "textarea" ? (
        <textarea value={value} onChange={update} />
      ) : type === "select" ? (
        <select value={value} onChange={update}>
          {options.map((o) => (
            <option key={o}>{o}</option>
          ))}
        </select>
      ) : (
        <input value={value} onChange={update} />
      )}
    </div>
  );
}

function exportJson(data) {
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `${data.codigo || "ficha-tecnica"}.json`;
  a.click();
}

function exportExcel(data) {
  const general = [
    ["Campo", "Información"],
    ["Nombre del Indicador", data.nombre],
    ["Código", data.codigo],
    ["Área Responsable", data.area],
    ["País", data.pais],
    ["Tipo", data.tipo],
    ["Categoría del Indicador", data.categoria],
    ["Categoría de Riesgo", data.riesgo],
    ["Fecha de Creación", data.fechaCreacion],
    ["Última Actualización", data.fechaActualizacion],
    ["Versión", data.version],
    ["Estado", data.estado],
    ["Marco Regulatorio", data.marco],
    ["Alcance / Jurisdicción", data.alcance],
  ];

  const tecnica = [
    ["Campo", "Información"],
    ["Objetivo", data.objetivo],
    ["Definición Conceptual", data.definicion],
    ["Interpretación", data.interpretacion],
    ["Unidad de Medida", data.unidad],
    ["Sentido de Interpretación", data.sentido],
    ["Fórmula", data.formula],
    ["Reglas de Negocio", data.reglas],
    ["Supuestos", data.supuestos],
  ];

  const variables = [["Variable", "Definición", "Fuente"]].concat(
    data.variables
      .split("\n")
      .filter(Boolean)
      .map((line) => line.split("|").map((x) => x.trim()))
  );

  const fuentes = [
    ["Campo", "Información"],
    ["Sistema Fuente", data.sistema],
    ["Tabla / Vista / Archivo", data.tabla],
    ["Responsable del Dato", data.responsableDato],
    ["Frecuencia de Actualización", data.freqAct],
    ["Frecuencia de Cálculo", data.freqCalc],
    ["Frecuencia de Monitoreo", data.freqMon],
    ["Frecuencia de Reporte", data.freqRep],
    ["Instancias de Reporte", data.instancias],
  ];

  const umbrales = [
    ["Estado", "Rango", "Interpretación / Acción"],
    ["Apetito", data.apetito, "Dentro del apetito de riesgo"],
    ["Tolerancia", data.tolerancia, data.accionTol],
    ["Capacidad", data.capacidad, data.accionCap],
    ["Metodología de Umbrales", data.metodoUmbrales, ""],
  ];

  const gobierno = [
    ["Rol / Campo", "Responsable / Resultado"],
    ["Dueño del Indicador", data.duenoIndicador],
    ["Dueño del Dato", data.duenoDato],
    ["Responsable Técnico", data.tecnico],
    ["Revisor", data.revisor],
    ["Aprobador", data.aprobador],
    ["Validación Conceptual", data.validacionConceptual],
    ["Validación Técnica", data.validacionTecnica],
    ["Validación Funcional", data.validacionFuncional],
    ["Supuestos y Limitaciones", data.limitaciones],
    ["Historial de Cambios", data.cambio],
    ["Elaboró", data.elaboro],
    ["Revisó", data.reviso],
    ["Aprobó", data.aprobo],
  ];

  const wb = XLSX.utils.book_new();
  [
    ["Información General", general],
    ["Definición Técnica", tecnica],
    ["Variables", variables],
    ["Fuentes y Frecuencia", fuentes],
    ["Umbrales", umbrales],
    ["Gobierno", gobierno],
  ].forEach(([name, rows]) => {
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = [{ wch: 28 }, { wch: 70 }, { wch: 45 }];
    XLSX.utils.book_append_sheet(wb, ws, name);
  });

  XLSX.writeFile(wb, `${data.codigo || "ficha-tecnica"}_ficha_tecnica.xlsx`);
}

function rows(text) {
  return text
    .split("\n")
    .filter(Boolean)
    .map((line, i) => {
      const [a, b, c] = line.split("|").map((x) => (x || "").trim());
      return (
        <tr key={i}>
          <td>{a}</td>
          <td>{b}</td>
          <td>{c}</td>
        </tr>
      );
    });
}

function Editor({ data, setData }) {
  return (
    <aside className="panel">
      <div className="panel-header">
        <h2>Constructor de Ficha Técnica</h2>
        <p>Completa los campos y revisa el documento en tiempo real.</p>
      </div>

      <div className="section">
        <details open>
          <summary>1. Información general <span>▾</span></summary>
          <div className="grid">
            <Field label="Nombre del indicador" name="nombre" value={data.nombre} setData={setData} />
            <Field label="Código" name="codigo" value={data.codigo} setData={setData} />
            <Field label="Área responsable" name="area" value={data.area} setData={setData} />
            <Field label="País" name="pais" value={data.pais} setData={setData} />
            <Field label="Tipo" name="tipo" value={data.tipo} setData={setData} type="select" options={["KRI", "KPI"]} />
            <Field label="Categoría del indicador" name="categoria" value={data.categoria} setData={setData} type="select" options={["Apetito de Riesgo", "Complementario", "Regulatorio", "Táctico u Operativo", "Apetito de Riesgo / Regulatorio"]} />
            <Field label="Categoría de riesgo" name="riesgo" value={data.riesgo} setData={setData} />
            <Field label="Estado" name="estado" value={data.estado} setData={setData} type="select" options={["Activo", "En revisión", "Descontinuado"]} />
            <Field label="Versión" name="version" value={data.version} setData={setData} />
            <Field label="Fecha actualización" name="fechaActualizacion" value={data.fechaActualizacion} setData={setData} />
            <Field label="Marco regulatorio" name="marco" value={data.marco} setData={setData} type="textarea" />
            <Field label="Alcance / jurisdicción" name="alcance" value={data.alcance} setData={setData} type="textarea" />
          </div>
        </details>
      </div>

      <div className="section">
        <details open>
          <summary>2. Definición y metodología <span>▾</span></summary>
          <div className="grid">
            <Field label="Objetivo" name="objetivo" value={data.objetivo} setData={setData} type="textarea" />
            <Field label="Definición conceptual" name="definicion" value={data.definicion} setData={setData} type="textarea" />
            <Field label="Interpretación" name="interpretacion" value={data.interpretacion} setData={setData} type="textarea" />
            <Field label="Unidad de medida" name="unidad" value={data.unidad} setData={setData} />
            <Field label="Sentido de interpretación" name="sentido" value={data.sentido} setData={setData} type="select" options={["Menor es mejor", "Mayor es mejor", "Dentro de rango es mejor"]} />
            <Field label="Fórmula" name="formula" value={data.formula} setData={setData} type="textarea" />
            <Field label="Reglas de negocio" name="reglas" value={data.reglas} setData={setData} type="textarea" />
            <Field label="Supuestos" name="supuestos" value={data.supuestos} setData={setData} type="textarea" />
          </div>
        </details>
      </div>

      <div className="section">
        <details>
          <summary>3. Fuentes, umbrales y gobierno <span>▾</span></summary>
          <div className="grid">
            <Field label="Variables | Definición | Fuente" name="variables" value={data.variables} setData={setData} type="textarea" />
            <Field label="Sistema fuente" name="sistema" value={data.sistema} setData={setData} />
            <Field label="Tabla / vista / archivo" name="tabla" value={data.tabla} setData={setData} />
            <Field label="Responsable del dato" name="responsableDato" value={data.responsableDato} setData={setData} />
            <Field label="Frecuencia actualización" name="freqAct" value={data.freqAct} setData={setData} />
            <Field label="Frecuencia cálculo" name="freqCalc" value={data.freqCalc} setData={setData} />
            <Field label="Frecuencia monitoreo" name="freqMon" value={data.freqMon} setData={setData} />
            <Field label="Frecuencia reporte" name="freqRep" value={data.freqRep} setData={setData} />
            <Field label="Instancias de reporte" name="instancias" value={data.instancias} setData={setData} />
            <Field label="Rango apetito" name="apetito" value={data.apetito} setData={setData} />
            <Field label="Rango tolerancia" name="tolerancia" value={data.tolerancia} setData={setData} />
            <Field label="Rango capacidad" name="capacidad" value={data.capacidad} setData={setData} />
            <Field label="Metodología de umbrales" name="metodoUmbrales" value={data.metodoUmbrales} setData={setData} type="textarea" />
            <Field label="Acción en tolerancia" name="accionTol" value={data.accionTol} setData={setData} type="textarea" />
            <Field label="Acción en capacidad" name="accionCap" value={data.accionCap} setData={setData} type="textarea" />
            <Field label="Dueño indicador" name="duenoIndicador" value={data.duenoIndicador} setData={setData} />
            <Field label="Responsable técnico" name="tecnico" value={data.tecnico} setData={setData} />
            <Field label="Revisor" name="revisor" value={data.revisor} setData={setData} />
            <Field label="Aprobador" name="aprobador" value={data.aprobador} setData={setData} />
          </div>
        </details>
      </div>

      <div className="actions">
        <button className="ghost" onClick={() => setData(initial)}><RotateCcw size={15}/> Ejemplo</button>
        <button className="ghost" onClick={() => exportJson(data)}><Database size={15}/> JSON</button>
        <button className="ghost" onClick={() => exportExcel(data)}><FileSpreadsheet size={15}/> Excel</button>
        <button className="primary" onClick={() => window.print()}><FileDown size={15}/> PDF</button>
      </div>
    </aside>
  );
}

function Preview({ data }) {
  return (
    <main className="preview">
      <div className="doc">
        <div className="doc-head">
          <div className="doc-logo"><img src={logo} alt="Grupo Ficohsa" /></div>
          <div className="codebox">{data.codigo || "Código"}<br /><span>Versión {data.version}</span></div>
        </div>

        <div className="titlebox">
          <h2>FICHA TÉCNICA DE INDICADOR DE RIESGO</h2>
          <p><b>Indicador:</b> {data.nombre}</p>
          <p><b>Código:</b> {data.codigo}</p>
          <p><b>Categoría:</b> {data.riesgo}</p>
          <p><b>Estado:</b> <span className="status ok">{data.estado}</span> <b className="ml">Versión:</b> {data.version}</p>
        </div>

        <h3>Resumen Ejecutivo</h3>
        <p>{data.objetivo}</p>

        <h3>Información General</h3>
        <table><tbody>
          {[
            ["Nombre del Indicador", data.nombre], ["Código", data.codigo], ["Área Responsable", data.area],
            ["País", data.pais], ["Tipo", data.tipo], ["Categoría del Indicador", data.categoria],
            ["Categoría de Riesgo", data.riesgo], ["Fecha de Creación", data.fechaCreacion],
            ["Última Actualización", data.fechaActualizacion], ["Versión", data.version],
            ["Estado", data.estado], ["Marco Regulatorio aplicable", data.marco], ["Alcance / Jurisdicción", data.alcance]
          ].map(([a, b]) => <tr key={a}><th>{a}</th><td>{b}</td></tr>)}
        </tbody></table>

        <h3>Objetivo</h3><p>{data.objetivo}</p>

        <h3>Definición Técnica</h3>
        <div className="chips">
          <span className="chip">{data.unidad}</span>
          <span className="chip">{data.sentido}</span>
          <span className="chip">{data.tipo}</span>
        </div>
        <p><b>Definición conceptual:</b> {data.definicion}</p>
        <p><b>Interpretación:</b> {data.interpretacion}</p>

        <h3>Metodología de Cálculo</h3>
        <p><b>Fórmula:</b> {data.formula}</p>
        <p><b>Reglas de negocio:</b> {data.reglas}</p>
        <p><b>Supuestos:</b> {data.supuestos}</p>

        <h3>Variables Utilizadas</h3>
        <table>
          <thead><tr><th>Variable</th><th>Definición</th><th>Fuente</th></tr></thead>
          <tbody>{rows(data.variables)}</tbody>
        </table>

        <h3>Fuentes de Información</h3>
        <table><tbody>
          <tr><th>Sistema Fuente</th><td>{data.sistema}</td></tr>
          <tr><th>Tabla / Vista / Archivo</th><td>{data.tabla}</td></tr>
          <tr><th>Responsable del Dato</th><td>{data.responsableDato}</td></tr>
          <tr><th>Frecuencia de Actualización</th><td>{data.freqAct}</td></tr>
        </tbody></table>

        <h3>Frecuencia</h3>
        <table><tbody>
          <tr><th>Frecuencia de Cálculo</th><td>{data.freqCalc}</td></tr>
          <tr><th>Frecuencia de Monitoreo</th><td>{data.freqMon}</td></tr>
          <tr><th>Frecuencia de Reporte</th><td>{data.freqRep}</td></tr>
          <tr><th>Instancias de Reporte</th><td>{data.instancias}</td></tr>
        </tbody></table>

        <h3>Umbrales y Límites</h3>
        <table>
          <thead><tr><th>Estado</th><th>Rango</th><th>Interpretación</th></tr></thead>
          <tbody>
            <tr><td><span className="status ok">Apetito</span></td><td>{data.apetito}</td><td>Dentro del apetito de riesgo</td></tr>
            <tr><td><span className="status warn">Tolerancia</span></td><td>{data.tolerancia}</td><td>Nivel de alerta</td></tr>
            <tr><td><span className="status bad">Capacidad</span></td><td>{data.capacidad}</td><td>Fuera del apetito de riesgo</td></tr>
          </tbody>
        </table>
        <p><b>Metodología de definición de umbrales:</b> {data.metodoUmbrales}</p>

        <h3>Acciones de Escalamiento</h3>
        <table>
          <thead><tr><th>Estado</th><th>Acción requerida</th><th>Responsable</th></tr></thead>
          <tbody>
            <tr><td>Tolerancia</td><td>{data.accionTol}</td><td>{data.area}</td></tr>
            <tr><td>Capacidad</td><td>{data.accionCap}</td><td>{data.area} / Alta Dirección</td></tr>
          </tbody>
        </table>

        <h3>Gobierno del Indicador</h3>
        <table><tbody>
          <tr><th>Dueño del Indicador</th><td>{data.duenoIndicador}</td></tr>
          <tr><th>Dueño del Dato</th><td>{data.duenoDato}</td></tr>
          <tr><th>Responsable Técnico</th><td>{data.tecnico}</td></tr>
          <tr><th>Revisor</th><td>{data.revisor}</td></tr>
          <tr><th>Aprobador</th><td>{data.aprobador}</td></tr>
        </tbody></table>

        <h3>Validación</h3>
        <table>
          <thead><tr><th>Tipo</th><th>Resultado</th><th>Fecha</th></tr></thead>
          <tbody>
            <tr><td>Conceptual</td><td>{data.validacionConceptual}</td><td>{data.fechaActualizacion}</td></tr>
            <tr><td>Técnica</td><td>{data.validacionTecnica}</td><td>{data.fechaActualizacion}</td></tr>
            <tr><td>Funcional</td><td>{data.validacionFuncional}</td><td>{data.fechaActualizacion}</td></tr>
          </tbody>
        </table>

        <h3>Supuestos y Limitaciones</h3><p>{data.limitaciones}</p>

        <h3>Historial de Cambios</h3>
        <table><thead><tr><th>Versión</th><th>Fecha</th><th>Descripción</th><th>Responsable</th></tr></thead>
          <tbody><tr><td>{data.version}</td><td>{data.fechaActualizacion}</td><td>{data.cambio}</td><td>{data.area}</td></tr></tbody></table>

        <h3>Aprobaciones</h3>
        <table><thead><tr><th>Rol</th><th>Nombre</th><th>Fecha</th><th>Firma</th></tr></thead>
          <tbody>
            <tr><td>Elaboró</td><td>{data.elaboro}</td><td></td><td></td></tr>
            <tr><td>Revisó</td><td>{data.reviso}</td><td></td><td></td></tr>
            <tr><td>Aprobó</td><td>{data.aprobo}</td><td></td><td></td></tr>
          </tbody></table>
      </div>
    </main>
  );
}

function App() {
  const [data, setData] = useState(initial);

  return (
    <div className="app">
      <header className="topbar">
        <div className="brand">
          <img className="logo-top" src={logo} alt="Grupo Ficohsa" />
          <div>
            <h1>Gobierno de Indicadores de Riesgo</h1>
            <p>Interfaz React para creación de fichas técnicas</p>
          </div>
        </div>
        <div className="badge">Uso interno · Grupo Ficohsa</div>
      </header>
      <div className="layout">
        <Editor data={data} setData={setData} />
        <Preview data={data} />
      </div>
    </div>
  );
}

createRoot(document.getElementById("root")).render(<App />);
