const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, HeadingLevel,
  BorderStyle, WidthType, ShadingType, VerticalAlign, PageNumber,
  TableOfContents,
} = require("docx");

const PAGE = { width: 12240, height: 15840 };
const MARGIN = { top: 1440, right: 1440, bottom: 1440, left: 1440 };

const border = { style: BorderStyle.SINGLE, size: 2, color: "CBD5E1" };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

function headerCell(text, width) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: "1F4E79", type: ShadingType.CLEAR },
    margins: cellMargins,
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ children: [new TextRun({ text, bold: true, color: "FFFFFF", size: 20 })] })],
  });
}

function bodyCell(text, width, opts = {}) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: opts.fill || "FFFFFF", type: ShadingType.CLEAR },
    margins: cellMargins,
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ alignment: opts.align || AlignmentType.LEFT, children: [new TextRun({ text, size: 20, bold: !!opts.bold })] })],
  });
}

function makeTable(headers, rows, widths, opts = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, children: headers.map((h, i) => headerCell(h, widths[i])) }),
      ...rows.map((r, idx) => new TableRow({
        children: r.map((c, i) => bodyCell(c, widths[i], {
          fill: idx % 2 === 0 ? "FFFFFF" : "F2F6FA",
          align: (opts.rightAlignCols || []).includes(i) ? AlignmentType.RIGHT : AlignmentType.LEFT,
        })),
      })),
    ],
  });
}

function h2(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 260, after: 120 }, children: [new TextRun(text)] });
}
function p(text, opts = {}) {
  return new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text, ...opts })] });
}
function bullet(text) {
  return new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 }, children: [new TextRun(text)] });
}
function numbered(text) {
  return new Paragraph({ numbering: { reference: "numbers", level: 0 }, spacing: { after: 60 }, children: [new TextRun(text)] });
}
function checkbullet(text) {
  return new Paragraph({ numbering: { reference: "checks", level: 0 }, spacing: { after: 60 }, children: [new TextRun(text)] });
}
function rule() {
  return new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "1F4E79", space: 1 } }, spacing: { after: 200 }, children: [] });
}

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Calibri", size: 22 } } },
    paragraphStyles: [
      { id: "Title", name: "Title", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 44, bold: true, color: "1F4E79", font: "Calibri" },
        paragraph: { spacing: { after: 80 } } },
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 30, bold: true, color: "1F4E79", font: "Calibri" },
        paragraph: { spacing: { before: 320, after: 160 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, color: "2E75B6", font: "Calibri" },
        paragraph: { spacing: { before: 260, after: 120 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, color: "444444", font: "Calibri" },
        paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 2 } },
    ],
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 620, hanging: 300 } } } }] },
      { reference: "numbers", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 620, hanging: 300 } } } }] },
      { reference: "checks", levels: [{ level: 0, format: LevelFormat.BULLET, text: "☐", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 620, hanging: 300 } } } }] },
    ],
  },
  sections: [{
    properties: { page: { size: PAGE, margin: MARGIN } },
    headers: {
      default: new Header({ children: [new Paragraph({
        alignment: AlignmentType.RIGHT,
        children: [new TextRun({ text: "Lectura Masiva de Contratos · Requerimientos y Validaciones", size: 16, color: "888888" })],
      })] }),
    },
    footers: {
      default: new Footer({ children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({ text: "Página ", size: 16, color: "888888" }),
          new TextRun({ children: [PageNumber.CURRENT], size: 16, color: "888888" }),
          new TextRun({ text: " de ", size: 16, color: "888888" }),
          new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, color: "888888" }),
        ],
      })] }),
    },
    children: [
      new Paragraph({ style: "Title", children: [new TextRun("Requerimientos y Validaciones del Proyecto")] }),
      new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "Lectura Masiva de Contratos — Solución sobre Microsoft Azure", size: 26, bold: true, color: "444444" })] }),
      new Paragraph({ spacing: { after: 300 }, children: [new TextRun({ text: "Documento preparado para revisión y validación del cliente · 22 de julio de 2026", size: 20, italics: true, color: "666666" })] }),
      rule(),

      new TableOfContents("Contenido", { hyperlink: true, headingStyleRange: "1-2" }),
      new Paragraph({ children: [new TextRun({ text: "" })], pageBreakBefore: true }),

      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("1. Resumen del proyecto")] }),
      p("El proyecto Lectura Masiva de Contratos tiene como objetivo implementar, sobre Microsoft Azure, una solución que procese contratos en formato PDF almacenados en SharePoint Online, extraiga la información relevante mediante Azure AI Document Intelligence, orqueste el flujo completo con Azure Logic Apps y prepare el envío seguro de las variables contractuales validadas hacia una plataforma externa alojada en Huawei Cloud, mediante una conexión VPN Site-to-Site (IPsec/IKEv2)."),
      p("El alcance de esta fase de despliegue considera únicamente los componentes de Microsoft Azure. No se desplegarán recursos en Huawei Cloud ni componentes adicionales fuera del diagrama aprobado por ambas partes."),

      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("2. Alcance de la solución")] }),
      makeTable(
        ["Área", "Alcance"],
        [
          ["Orquestación", "Azure Logic Apps coordina lectura, extracción, validación y envío"],
          ["Fuente documental", "SharePoint Online como repositorio de contratos PDF"],
          ["Acceso documental", "Conector SharePoint y/o Microsoft Graph"],
          ["Extracción OCR", "Azure AI Document Intelligence"],
          ["Red", "Azure Virtual Network, subredes y VPN Gateway"],
          ["Conectividad externa", "VPN Site-to-Site hacia Huawei Cloud"],
          ["Identidad", "Managed Identity y permisos Microsoft Entra ID"],
          ["Gobierno", "Azure Policy, etiquetas (tags), regiones permitidas y RBAC"],
          ["Infraestructura como código", "Bicep, Azure CLI y PowerShell"],
        ],
        [2500, 6860]
      ),

      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("3. Componentes Azure requeridos")] }),
      makeTable(
        ["Componente", "Propósito", "Requerido"],
        [
          ["Resource Group de red", "Contener los recursos de red", "Sí"],
          ["Resource Group de integración", "Contener Logic Apps y Document Intelligence", "Sí"],
          ["Resource Group de seguridad", "Contener políticas y asignaciones RBAC", "Sí"],
          ["Azure Virtual Network", "Red privada de la solución", "Sí"],
          ["GatewaySubnet", "Subred obligatoria para el VPN Gateway", "Sí"],
          ["Subred de integración", "Subred para Logic Apps Standard, si aplica", "Sí"],
          ["Public IP", "IP pública del VPN Gateway", "Sí"],
          ["Azure VPN Gateway", "Conectividad Site-to-Site IPsec/IKEv2", "Sí"],
          ["Local Network Gateway", "Representa el gateway remoto de Huawei", "Sí"],
          ["VPN Connection", "Conecta el VPN Gateway de Azure con el de Huawei", "Sí"],
          ["Azure Logic Apps", "Orquestador del flujo documental", "Sí"],
          ["Azure AI Document Intelligence", "OCR y extracción de información contractual", "Sí"],
          ["Managed Identity", "Identidad segura para la ejecución del flujo", "Sí"],
          ["API Connection SharePoint", "Conector para leer contratos desde SharePoint Online", "Sí, si aplica"],
          ["Azure Policy", "Restringe regiones, exige etiquetas y limita recursos", "Sí"],
          ["RBAC Role Assignments", "Asigna los permisos mínimos requeridos", "Sí"],
        ],
        [3200, 4960, 1200]
      ),
      p("Nota: si Document Intelligence se despliega como cuenta de Microsoft Foundry (AI Services multi-servicio con proyectos), el resource group de integración también es el alcance recomendado para las asignaciones de rol Foundry (ver sección 11.1).", { italics: true, size: 20, color: "555555" }),

      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("4. Componentes fuera de alcance")] }),
      makeTable(
        ["Componente", "Motivo"],
        [
          ["PostgreSQL en Huawei", "Es parte de Huawei Cloud, no de Azure"],
          ["OBS en Huawei", "Es parte de Huawei Cloud, no de Azure"],
          ["Aplicación de validación (Angular)", "Fuera de la sección Azure del diagrama"],
          ["Módulo externo (Python)", "Fuera de la sección Azure del diagrama"],
          ["AKS", "No aparece en el diagrama aprobado"],
          ["Azure SQL / Cosmos DB", "No aparecen en la sección Azure"],
          ["Service Bus / Event Grid", "No aparecen en la sección Azure"],
          ["Key Vault", "Se incluirá solo si el cliente lo aprueba como dependencia técnica"],
          ["Application Insights", "Se incluirá solo si el cliente aprueba observabilidad adicional"],
        ],
        [3200, 6160]
      ),
      p("Nota de seguridad: aunque Key Vault queda fuera del alcance inicial, la llave compartida (shared key) de la VPN y cualquier otro secreto no se almacenarán en texto plano. Mientras Key Vault no sea aprobado, estos valores se inyectarán en el despliegue mediante parámetros seguros (@secure()), sin quedar versionados en el repositorio.", { italics: true, size: 20, color: "555555" }),

      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("5. Diagrama de despliegue (resumen)")] }),
      p("El flujo conecta SharePoint Online (Microsoft 365) con un conjunto de recursos de Azure organizados en tres grupos de recursos:"),
      bullet("rg-lectura-contratos-network-<env>: VNet, GatewaySubnet, VPN Gateway, Public IP, Local Network Gateway y VPN Connection hacia Huawei."),
      bullet("rg-lectura-contratos-integration-<env>: Azure Logic Apps, Managed Identity, API Connection SharePoint y Azure AI Document Intelligence."),
      bullet("rg-lectura-contratos-security-<env>: Azure Policy (regiones, tags, tipos de recursos permitidos) y asignaciones RBAC."),
      p("Huawei Cloud (VPN Gateway, PostgreSQL, OBS) permanece fuera del alcance de despliegue de Azure y solo se conecta mediante la VPN Site-to-Site. El diagrama completo (Mermaid y textual) está disponible en el README técnico del repositorio del proyecto."),

      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("6. Flujo funcional esperado")] }),
      numbered("Los contratos PDF se almacenan en SharePoint Online."),
      numbered("Azure Logic Apps detecta o consulta los documentos disponibles."),
      numbered("Logic Apps obtiene el archivo mediante el conector de SharePoint o Microsoft Graph."),
      numbered("El PDF se envía a Azure AI Document Intelligence."),
      numbered("Document Intelligence extrae texto, estructura, tablas y campos."),
      numbered("Logic Apps evalúa las variables contractuales y su nivel de confianza."),
      numbered("Las variables con confianza suficiente se preparan para el envío."),
      numbered("Las variables con baja confianza se marcan para revisión humana."),
      numbered("El envío hacia Huawei se realiza a través de la VPN Site-to-Site."),

      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("7. Regiones permitidas")] }),
      p("Todos los recursos de Azure se desplegarán únicamente en regiones del continente americano. La región recomendada para iniciar la demo es East US (eastus)."),
      makeTable(
        ["Prioridad", "Región", "Código Azure", "Uso recomendado"],
        [
          ["1", "East US", "eastus", "Región por defecto para la demo"],
          ["2", "East US 2", "eastus2", "Alternativa principal"],
          ["3", "South Central US", "southcentralus", "Alternativa cercana a México/LATAM"],
          ["4", "Central US", "centralus", "Alternativa para pruebas"],
          ["5", "West US 2", "westus2", "Alternativa oeste de Estados Unidos"],
          ["6", "Canada Central", "canadacentral", "Alternativa Canadá"],
          ["7", "Brazil South", "brazilsouth", "Alternativa LATAM, requiere validación de SKUs"],
          ["8", "Mexico Central", "mexicocentral", "Alternativa México, requiere validación completa"],
        ],
        [1100, 2200, 2200, 3860]
      ),

      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("8. Requerimientos de red")] }),
      makeTable(
        ["Requerimiento", "Valor recomendado / origen"],
        [
          ["Tipo de conectividad", "VPN Site-to-Site IPsec/IKEv2"],
          ["CIDR Huawei", "A proporcionar por el cliente"],
          ["IP pública del VPN Huawei", "A proporcionar por el cliente"],
          ["Llave compartida (shared key) VPN", "A proporcionar por el cliente, por canal seguro"],
          ["No traslape de rangos IP", "Azure, Huawei y red corporativa no deben compartir rangos"],
          ["Tamaño mínimo de GatewaySubnet", "/27 como mínimo; /26 (usado en dev/test/prod) da margen para SKUs mayores o activo-activo"],
        ],
        [3600, 5760]
      ),

      h2("8.1 Detalle de rangos IP y direcciones disponibles"),
      makeTable(
        ["Requerimiento", "CIDR", "Direcciones totales", "IPs utilizables"],
        [
          ["Rango IP Azure dev", "10.240.0.0/20", "4096", "4096 (bloque padre; ver subredes)"],
          ["GatewaySubnet dev", "10.240.0.0/26", "64", "59"],
          ["Logic Apps subnet dev", "10.240.1.0/26", "64", "59"],
          ["Rango IP Azure test", "10.240.16.0/20", "4096", "4096 (bloque padre; ver subredes)"],
          ["GatewaySubnet test", "10.240.16.0/26", "64", "59"],
          ["Logic Apps subnet test", "10.240.17.0/26", "64", "59"],
          ["Rango IP Azure prod", "10.240.32.0/20", "4096", "4096 (bloque padre; ver subredes)"],
          ["GatewaySubnet prod", "10.240.32.0/26", "64", "59"],
          ["Logic Apps subnet prod", "10.240.33.0/26", "64", "59"],
        ],
        [2600, 2400, 2000, 2360],
        { rightAlignCols: [2, 3] }
      ),
      p("Nota: \"Direcciones totales\" es el tamaño matemático del bloque (2 elevado a 32 menos el prefijo). \"IPs utilizables\" en las subredes /26 descuenta las 5 direcciones que Azure reserva por subred (dirección de red, gateway, dos reservadas y broadcast). Los rangos /20 son bloques padre reservados para alojar las subredes de cada ambiente; su disponibilidad real depende de cómo se subdividan.", { italics: true, size: 20, color: "555555" }),

      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("9. Requerimientos de Microsoft 365")] }),
      makeTable(
        ["Requerimiento", "Detalle"],
        [
          ["SharePoint Online", "Sitio y biblioteca de documentos para contratos PDF"],
          ["Carpeta origen", "Ruta donde se colocarán los contratos"],
          ["Documentos de prueba", "Contratos dummy, sin información sensible"],
          ["Microsoft Graph", "Acceso para leer archivos y metadatos"],
          ["Permiso recomendado", "Sites.Selected"],
          ["Permiso alternativo (demo)", "Sites.Read.All"],
          ["Consentimiento de administrador", "Requerido si se usan permisos de tipo aplicación"],
        ],
        [3600, 5760]
      ),

      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("10. Requerimientos de Azure Landing Zone")] }),
      makeTable(
        ["Requerimiento", "Descripción"],
        [
          ["Subscription", "Suscripción dedicada o controlada para la demo"],
          ["Resource groups", "Separación por red, integración y seguridad"],
          ["Naming convention", "Nombres por workload, ambiente y región"],
          ["Tags obligatorios", "environment, owner, costCenter, workload, dataClassification"],
          ["Azure Policy", "Regiones permitidas, tags requeridos y tipos de recursos permitidos"],
          ["RBAC", "Acceso por roles mínimos necesarios"],
          ["Presupuesto", "Recomendable definir budget o alerta de costos"],
          ["IaC", "Todo despliegue debe realizarse mediante Bicep"],
        ],
        [3200, 6160]
      ),

      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("11. Niveles de permisos y roles")] }),
      h2("Roles para despliegue inicial"),
      makeTable(
        ["Rol", "Alcance recomendado", "Uso"],
        [
          ["Owner", "Temporal a nivel subscription, si se requiere", "Delegar permisos iniciales y configurar RBAC"],
          ["Contributor", "Subscription o resource groups", "Crear recursos Azure"],
          ["Network Contributor", "Resource group de red", "Crear VNet, subredes, VPN Gateway, Local Network Gateway y conexiones"],
          ["User Access Administrator", "Subscription o resource groups", "Asignar roles a Managed Identity y operadores"],
          ["Cognitive Services Contributor", "Resource group de integración", "Crear Document Intelligence"],
          ["Logic App Contributor", "Resource group de integración", "Crear y administrar Logic Apps"],
          ["Reader", "Subscription o resource groups", "Validación, auditoría y revisión"],
        ],
        [2800, 3400, 3160]
      ),
      h2("Roles para operación posterior"),
      makeTable(
        ["Equipo", "Rol sugerido", "Alcance"],
        [
          ["Equipo de red", "Network Contributor", "rg-lectura-contratos-network-<env>"],
          ["Equipo de integración", "Logic App Contributor", "rg-lectura-contratos-integration-<env>"],
          ["Equipo de AI/documentos", "Cognitive Services User / Contributor según operación", "Document Intelligence"],
          ["Equipo de seguridad", "Security Reader / Reader", "Subscription o resource groups"],
          ["Equipo de plataforma", "Contributor limitado", "Resource groups del proyecto"],
        ],
        [2800, 3400, 3160]
      ),

      h2("11.1 RBAC de Microsoft Foundry para el recurso de IA"),
      p("Si Azure AI Document Intelligence se despliega como parte de una cuenta de Microsoft Foundry (AI Services multi-servicio con proyectos), aplican los siguientes conceptos de RBAC, adicionales a los roles clásicos de Azure. Referencia: Role-based access control for Microsoft Foundry (Microsoft Learn)."),
      p("Alcances (scopes) de asignación de roles en Foundry:", { bold: true, size: 20 }),
      makeTable(
        ["Alcance", "Descripción"],
        [
          ["Subscription", "Alcance más amplio; normalmente reservado para el equipo de IT/Landing Zone"],
          ["Resource Group", "rg-lectura-contratos-integration-<env> — grupo de recursos que contendrá la cuenta Foundry/AI Services"],
          ["Foundry resource (cuenta)", "Límite administrativo, de seguridad y monitoreo del entorno Foundry"],
          ["Foundry project", "Sub-alcance dentro del recurso Foundry para organizar el trabajo y controlar el acceso a APIs y herramientas"],
          ["Agente individual", "Alcance más granular; controla el acceso al endpoint de un agente específico"],
        ],
        [3200, 6160]
      ),
      p("Roles integrados (built-in) de Foundry recomendados:", { bold: true, size: 20 }),
      makeTable(
        ["Rol", "Descripción", "Uso recomendado en este proyecto"],
        [
          ["Foundry User", "Acceso de lectura al proyecto/recurso y permisos de construcción y prueba", "Equipo de AI/documentos que configura y prueba la extracción"],
          ["Foundry Project Manager", "Gestiona proyectos Foundry; puede construir y asignar Foundry User a otros", "Líder técnico del equipo de integración/AI"],
          ["Foundry Account Owner", "Administra la cuenta completa: despliega modelos, audita conexiones y cómputo; puede asignar Foundry User", "Responsable de la solución en el resource group de integración"],
          ["Foundry Owner", "Control total (gestión + construcción); puede asignar Foundry User, ACR y roles de monitoreo", "Solo para configuración inicial o soporte crítico"],
          ["Foundry Agent Consumer", "Mínimo privilegio: solo invoca el endpoint de un agente, sin crearlo/modificarlo", "Sistemas o cuentas de servicio que solo consumen el resultado (integración con Huawei)"],
        ],
        [2400, 4160, 2800]
      ),
      p("Importante: no usar los roles que inician con \"Cognitive Services\" (p. ej. Cognitive Services Contributor) para escenarios de Foundry; están pensados para acceso directo a recursos de AI Services, no al modelo de proyectos de Foundry. Tampoco usar \"Azure AI Developer\" para Foundry, ya que aplica a Machine Learning workspaces/hubs, no a proyectos Foundry.", { italics: true, size: 20, color: "B45309" }),
      p("Asignaciones mínimas para iniciar:", { bold: true, size: 20 }),
      numbered("Asignar el rol Foundry User al principal de usuario que configurará el flujo (en el recurso Foundry o en el resource group de integración)."),
      numbered("Asignar el rol Foundry User a la Managed Identity de Azure Logic Apps para que pueda invocar Document Intelligence/Foundry en tiempo de ejecución."),
      numbered("Si el usuario que crea el proyecto ya tiene rol Owner a nivel de subscription o resource group, ambas asignaciones pueden completarse automáticamente al crear el proyecto."),
      p("Mapeo sugerido de RBAC empresarial para este proyecto:", { bold: true, size: 20 }),
      makeTable(
        ["Persona / Identidad", "Rol y alcance", "Propósito"],
        [
          ["IT admin / Landing Zone", "Owner en subscription", "Garantiza estándares empresariales y delega roles de administración de Foundry"],
          ["Responsable de integración", "Foundry Account Owner en el recurso Foundry (rg-integration)", "Despliega y audita el recurso Foundry/AI Services, asigna Foundry User a su equipo"],
          ["Desarrollador del flujo (Logic Apps)", "Foundry User en el proyecto Foundry", "Construye y prueba la extracción de variables contractuales"],
          ["Managed Identity de Logic Apps", "Foundry User en el recurso/proyecto Foundry", "Permite la invocación de Document Intelligence en tiempo de ejecución sin usar claves"],
          ["Integración con Huawei (si aplica)", "Foundry Agent Consumer a nivel de proyecto o agente", "Consumo del resultado final sin permisos de construcción"],
        ],
        [2400, 3400, 3560]
      ),
      p("Notas y limitaciones relevantes:", { bold: true, size: 20 }),
      bullet("Se recomienda autenticación con Microsoft Entra ID (RBAC) en lugar de claves; con autenticación por clave se obtiene acceso completo sin restricciones de rol."),
      bullet("Se requiere el rol Contributor a nivel subscription para ver y purgar cuentas Foundry eliminadas."),
      bullet("Los usuarios con rol Contributor pueden desplegar modelos en Foundry."),
      bullet("Para crear roles personalizados sobre el recurso, se requiere el rol Owner en dicho recurso."),
      bullet("Si se necesita un rol personalizado más restrictivo que Foundry User, se puede definir un rol custom a nivel de subscription."),

      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("12. Parámetros requeridos para el despliegue")] }),
      makeTable(
        ["Parámetro", "Descripción", "Ejemplo"],
        [
          ["environment", "Ambiente de despliegue", "dev, test, prod"],
          ["location", "Región Azure permitida", "eastus"],
          ["workloadName", "Nombre del workload", "lectura-contratos"],
          ["azureVnetAddressSpace", "CIDR de la VNet Azure", "10.240.0.0/20"],
          ["gatewaySubnetPrefix", "CIDR de GatewaySubnet", "10.240.0.0/26"],
          ["logicAppsSubnetPrefix", "CIDR de subred Logic Apps", "10.240.1.0/26"],
          ["huaweiVpnPublicIp", "IP pública del gateway Huawei", "Proporcionado por cliente"],
          ["huaweiAddressPrefixes", "CIDR remoto Huawei", "Proporcionado por cliente"],
          ["vpnSharedKey", "Llave compartida VPN", "Valor seguro"],
          ["documentIntelligenceSku", "SKU del servicio OCR", "S0 o definido por cliente"],
          ["logicAppSku", "SKU de Logic Apps", "Según decisión de arquitectura"],
          ["allowedLocations", "Regiones permitidas", "Lista de regiones de América"],
        ],
        [2600, 3800, 2960]
      ),

      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("13. Validaciones previas al despliegue")] }),
      p("Antes de iniciar el despliegue se validará lo siguiente:"),
      checkbullet("La región seleccionada pertenece al continente americano."),
      checkbullet("Todos los servicios requeridos están disponibles en la región elegida."),
      checkbullet("Los resource providers de Azure están registrados en la suscripción."),
      checkbullet("Existen cuotas suficientes para VPN Gateway, IP pública y Azure AI Services."),
      checkbullet("Los rangos IP de Azure no se traslapan con Huawei ni con la red corporativa."),
      checkbullet("La IP pública del gateway Huawei fue proporcionada."),
      checkbullet("La llave compartida (shared key) de la VPN fue entregada por un canal seguro."),
      checkbullet("Existen permisos suficientes para crear recursos y asignar roles (RBAC)."),
      checkbullet("SharePoint Online cuenta con la biblioteca y los contratos PDF de prueba."),
      checkbullet("Los permisos de Microsoft Graph fueron aprobados por el administrador."),
      checkbullet("Se confirmó el modelo de despliegue de Document Intelligence (independiente vs. Microsoft Foundry) y las asignaciones de rol correspondientes."),

      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("14. Entregables esperados")] }),
      makeTable(
        ["Entregable", "Descripción"],
        [
          ["Infraestructura como código", "Bicep modular para red, VPN, Logic Apps, AI y políticas"],
          ["Scripts de despliegue", "PowerShell para validación, what-if y despliegue"],
          ["Documentación técnica", "Arquitectura, prerequisitos, red, seguridad y validación regional"],
          ["Diagramas", "Diagramas Mermaid de arquitectura, red y despliegue"],
          ["README", "Guía de ejecución para validar y desplegar la solución"],
        ],
        [3200, 6160]
      ),

      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("15. Consideraciones importantes")] }),
      bullet("La demo no desplegará recursos en Huawei Cloud."),
      bullet("La conexión hacia Huawei se modelará desde Azure usando Local Network Gateway y VPN Connection."),
      bullet("Se deben usar únicamente regiones Azure del continente americano."),
      bullet("No se deben guardar secretos, contratos reales ni datos sensibles en el repositorio."),
      bullet("Cualquier componente adicional fuera del diagrama debe ser aprobado explícitamente por el cliente."),
      bullet("Si se usa Logic Apps Standard, podrían requerirse dependencias técnicas de plataforma; estas deben documentarse claramente antes del despliegue."),
      bullet("Sin Application Insights aprobado, se recomienda al menos habilitar los logs de diagnóstico nativos de Logic Apps y Document Intelligence hacia Log Analytics para auditar fallos y confianza de extracción."),
      bullet("La llave compartida (vpnSharedKey) y cualquier credencial deben tratarse como secretos incluso mientras Key Vault no esté aprobado; usar parámetros @secure() en Bicep y secretos de pipeline para su inyección."),
      bullet("Si Document Intelligence se despliega como cuenta de Microsoft Foundry, se debe seguir el modelo de RBAC descrito en la sección 11.1, evitando roles de Cognitive Services o Azure AI Developer para estos escenarios."),

      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("16. Información y accesos requeridos del cliente")] }),
      p("Para iniciar el despliegue, se solicita al cliente proporcionar o confirmar lo siguiente:"),
      bullet("Suscripción de Azure dedicada o controlada para el proyecto, con presupuesto o alerta de costos definida."),
      bullet("Acceso con rol Owner temporal (o equivalente) para la configuración inicial de red, identidad y RBAC."),
      bullet("Confirmación de la región Azure a utilizar (recomendado: East US)."),
      bullet("Datos de conectividad Huawei: CIDR remoto, IP pública del VPN Gateway y llave compartida (por canal seguro)."),
      bullet("Sitio y biblioteca de SharePoint Online con los contratos de prueba, y aprobación de los permisos de Microsoft Graph."),
      bullet("Convenciones de nombres, etiquetas (tags) obligatorias y políticas de gobierno a respetar (si el cliente tiene estándares propios)."),
      bullet("Confirmación sobre la inclusión de Key Vault y/o Application Insights como dependencias adicionales aprobadas."),
      bullet("Confirmación del modelo de despliegue de Document Intelligence (recurso Cognitive Services independiente vs. cuenta de Microsoft Foundry) para definir el modelo de RBAC aplicable (sección 11.1)."),

      new Paragraph({ spacing: { before: 400 }, border: { top: { style: BorderStyle.SINGLE, size: 6, color: "1F4E79", space: 8 } }, children: [
        new TextRun({ text: "Quedamos atentos a la confirmación de la información solicitada en la sección 16 para continuar con la planeación del despliegue.", italics: true, size: 20, color: "444444" }),
      ] }),
    ],
  }],
});

Packer.toBuffer(doc).then((buffer) => {
  const outPath = path.join(__dirname, "..", "Requerimientos_y_Validaciones_Cliente.docx");
  fs.writeFileSync(outPath, buffer);
  console.log("Generado:", outPath);
});
