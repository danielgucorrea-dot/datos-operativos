const { default: makeWASocket, useMultiFileAuthState, fetchLatestBaileysVersion, downloadMediaMessage } = require('baileys')
const qrcode = require('qrcode-terminal')
const fs = require('fs')
const path = require('path')
const https = require('https')
const { Ollama } = require('ollama')
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, BorderStyle, WidthType } = require('docx')

const ollama = new Ollama({ host: 'http://localhost:11434' })

// ===== CONFIGURACIÓN =====
const GRUPO_ASESORES = '120363252943114933@g.us'
const MODO_PRUEBA = true
const NUMERO_PRUEBA = '5493814461809'
const JID_PRUEBA = '60886144299236'
const NUMEROS_PRUEBA_EXTRA = ['5493815266329', '5493812081709', '5493814167742', '5493815882162', '5493825402205', '5493816318179']
const NUMERO_MULTAS = '5493812081709'
const NOMBRE_GRUPO_BAM = 'BAM'
let GRUPO_BAM_ID = null
const NUMERO_ORDEN_SERVICIO = '5493815391601' // Número que recibe la orden de servicio diaria

// Números de personal interno — excluidos de base electoral y órdenes de servicio
// (salvo mensajes que envíen al grupo BAM)
const NUMEROS_PERSONAL = [
    '3814440012', '3813600157', '3816069950',
    '3816669333', '3816940661', '3816684407', '3816240722',
    '3815391601'
].map(n => n.replace(/\D/g, '')) // solo dígitos

// Números de prueba con restricción horaria (responden 13hs-8hs, ignorados 8-13hs L-V)
const NUMEROS_PRUEBA_HORARIO = ['5493812081709', '5493815266329', '5493814167742', '5493815882162', '5493825402205', '5493816318179']
// JIDs @lid conocidos de los números de prueba (para cuando no están en lid_map.json)
const JIDS_PRUEBA_CONOCIDOS = {
    '5493815266329': '', // se completa cuando llega el primer mensaje
    '5493812081709': '',
    '5493814167742': '',
    '5493815882162': ''
}

function esPruebaHorarioActivo(jid) {
    // Resolver el número real: el JID puede ser @lid (sin número visible)
    // Buscar en lid_map, o extraer directo si es @s.whatsapp.net
    const pnResuelto = buscarPNenMap(jid) || jid.replace(/@.*$/, '')
    const pnLimpio = pnResuelto.replace(/\D/g, '').replace(/^549/, '')

    // Verificar si el número resuelto está en la lista de prueba con horario
    const esPruebaHorario = NUMEROS_PRUEBA_HORARIO.some(n => {
        const nLimpio = n.replace(/\D/g, '').replace(/^549/, '')
        return pnLimpio === nLimpio || pnLimpio.endsWith(nLimpio) || nLimpio.endsWith(pnLimpio)
    })
    if (!esPruebaHorario) return false

    const ahora = new Date()
    const hora = ahora.getHours()
    const dia = ahora.getDay() // 0=dom, 6=sab
    const esLunesAViernes = dia >= 1 && dia <= 5
    // De 8 a 13 lunes a viernes → NO responder (horario operativo real)
    if (esLunesAViernes && hora >= 8 && hora < 13) return false
    return true
}

function esPersonalInterno(numero) {
    const limpio = numero.replace(/\D/g, '').replace(/^549/, '').replace(/^54/, '')
    return NUMEROS_PERSONAL.some(p => limpio.endsWith(p) || p.endsWith(limpio))
}

// Validación básica de que el texto parece una dirección
function pareceDireccion(texto) {
    if (!texto) return false
    const t = texto.trim()
    if (t.length < 5) return false
    // Debe contener al menos un número, o palabras típicas de dirección
    const tieneNumero = /\d/.test(t)
    const palabrasDireccion = /(calle|avenida|av\.|pasaje|barrio|b°|esquina|esq\.|y\s|entre\s)/i.test(t)
    // Rechazar saludos comunes
    const saludos = /^(hola|buenas|buen[oa]s|que tal|hello|hi|gracias|ok|si|no|listo)$/i
    if (saludos.test(t)) return false
    return tieneNumero || palabrasDireccion
}

// ===== FERIADOS ARGENTINA 2026 =====
const FERIADOS = [
    '2026-01-01','2026-02-16','2026-02-17','2026-03-23','2026-03-24',
    '2026-04-02','2026-04-03','2026-04-06','2026-04-07','2026-04-08',
    '2026-04-09','2026-04-10','2026-05-01','2026-05-25','2026-06-15',
    '2026-06-20','2026-07-09','2026-08-17','2026-10-12','2026-11-20',
    '2026-12-08','2026-12-25'
]

// ===== LÍMITE NORTE/SUR (desde CSV) =====
// Polilínea que divide zona Norte (azul) de zona Sur (naranja)
const LIMITE_ZONA = [
    [-65.1765382, -26.836533],
    [-65.2172649, -26.8285294],
    [-65.2140033, -26.8164081],
    [-65.2078279, -26.8175389],
    [-65.2076563, -26.8172007],
    [-65.206798,  -26.8134473],
    [-65.2234062, -26.7932608],
    [-65.2237327, -26.7937937],
    [-65.2316291, -26.8250485],
    [-65.23223,   -26.8252783],
    [-65.2647598, -26.8183846]
]

// ===== GEOCODIFICACIÓN =====
function geocodificarQuery(query) {
    return new Promise((resolve) => {
        const encoded = encodeURIComponent(query)
        const options = {
            hostname: 'nominatim.openstreetmap.org',
            path: `/search?q=${encoded}&format=json&limit=1&countrycodes=ar`,
            method: 'GET',
            headers: { 'User-Agent': 'BotAmbienteSMT/1.0 (contacto@smt.gob.ar)' }
        }
        const req = https.request(options, (res) => {
            let data = ''
            res.on('data', chunk => data += chunk)
            res.on('end', () => {
                try {
                    const json = JSON.parse(data)
                    if (json.length > 0) {
                        resolve({ lat: parseFloat(json[0].lat), lon: parseFloat(json[0].lon), display: json[0].display_name })
                    } else {
                        resolve(null)
                    }
                } catch (e) {
                    resolve(null)
                }
            })
        })
        req.on('error', () => resolve(null))
        req.end()
    })
}

// Limpia el texto de dirección quitando palabras irrelevantes para geocodificar
function limpiarDireccionParaGeo(direccion) {
    return direccion
        .replace(/\bal frente del?\b/gi, '')
        .replace(/\bdesde\b.*$/gi, '') // cortar "desde X hasta Y" -> dejar solo el inicio
        .replace(/\bavenida\b/gi, 'Av')
        .replace(/[°ª]/g, '')
        .replace(/\s{2,}/g, ' ')
        .trim()
}

// Intenta geocodificar probando variantes de la dirección
async function geocodificar(direccion) {
    const limpio = limpiarDireccionParaGeo(direccion)

    const intentos = [
        `${limpio}, San Miguel de Tucuman, Tucuman, Argentina`,
        `${limpio}, San Miguel de Tucuman, Argentina`,
        `${limpio}, Tucuman, Argentina`
    ]

    for (const query of intentos) {
        const resultado = await geocodificarQuery(query)
        if (resultado) return resultado
        await new Promise(r => setTimeout(r, 1100)) // respetar rate limit Nominatim (1 req/seg)
    }
    return null
}

// Determina de qué lado de la polilínea cae el punto
// Retorna 'norte' o 'sur' comparando la latitud con el segmento más cercano
function detectarZonaPorCoordenadas(lat, lon) {
    let minDist = Infinity
    let zonaDetectada = 'sur'

    for (let i = 0; i < LIMITE_ZONA.length - 1; i++) {
        const [x1, y1] = LIMITE_ZONA[i]
        const [x2, y2] = LIMITE_ZONA[i + 1]

        const dx = x2 - x1
        const dy = y2 - y1
        const t = Math.max(0, Math.min(1, ((lon - x1) * dx + (lat - y1) * dy) / (dx * dx + dy * dy)))
        const projX = x1 + t * dx
        const projY = y1 + t * dy
        const dist = Math.sqrt((lon - projX) ** 2 + (lat - projY) ** 2)

        if (dist < minDist) {
            minDist = dist
            zonaDetectada = lat > projY ? 'norte' : 'sur'
        }
    }
    return zonaDetectada
}

// ===== DICCIONARIO DE CALLES -> ZONA (aprendizaje automático) =====
function cargarCallesZona() {
    if (fs.existsSync('calles_zona.json')) return JSON.parse(fs.readFileSync('calles_zona.json', 'utf8'))
    return {}
}
function guardarCallesZona(dic) {
    fs.writeFileSync('calles_zona.json', JSON.stringify(dic, null, 2))
}

function normalizarNombreCalle(nombre) {
    return nombre
        .toLowerCase()
        .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // quitar acentos
        // Abreviaturas y prefijos de tipo de vía
        .replace(/\b(av|avda|avenida|calle|cl|pje|psj|pasaje|diag|diagonal|bv|boulevard|ruta|rto|cto|coronel|col|gral|general|pdte|presidente|dr|ing|sgto|sargento|pte|tte|teniente)\b\.?/g, '')
        // Correcciones ortográficas comunes en SMT
        .replace(/\bjulio\b/g, 'julio')
        .replace(/\bvillaroel\b/g, 'villarroel')
        .replace(/\bpalacios\b/g, 'palacios')
        .replace(/\bnoreste\b/g, 'noreste')
        .replace(/\bsuroeste\b/g, 'suroeste')
        .replace(/\bnogues\b/g, 'nogues')
        .replace(/\bnougues\b/g, 'nogues')
        .replace(/[^a-z0-9\s]/g, '')
        .replace(/\s+/g, ' ')
        .trim()
}

// Similitud entre dos strings (para fuzzy matching de calles)
function similitudCalles(a, b) {
    if (a === b) return 1
    if (a.includes(b) || b.includes(a)) return 0.9
    // Contar palabras en común
    const wordsA = a.split(' ').filter(w => w.length > 2)
    const wordsB = b.split(' ').filter(w => w.length > 2)
    if (!wordsA.length || !wordsB.length) return 0
    const comunes = wordsA.filter(w => wordsB.some(wb => wb.includes(w) || w.includes(wb)))
    return comunes.length / Math.max(wordsA.length, wordsB.length)
}

// Extrae posibles nombres de calle de una dirección tipo "Calle X y Calle Y" o "Calle X 123"
function extraerNombresCalles(direccion) {
    const nombres = []
    // Separar por conectores comunes
    const partes = direccion.split(/\s+(?:y|esq\.?|esquina|entre|desde|hasta|al frente del?|bajo del?)\s+/i)
    for (let parte of partes) {
        // Quitar números sueltos al final/inicio
        parte = parte.replace(/\b\d+\b/g, '').trim()
        const norm = normalizarNombreCalle(parte)
        if (norm.length >= 3) nombres.push(norm)
    }
    return nombres
}

function registrarCallesZona(direccion, zona) {
    const dic = cargarCallesZona()
    const nombres = extraerNombresCalles(direccion)
    let cambios = false
    for (const nombre of nombres) {
        if (!dic[nombre]) {
            dic[nombre] = zona
            cambios = true
        }
    }
    if (cambios) guardarCallesZona(dic)
}

// Busca si alguna calle de la dirección ya está en el diccionario
function buscarZonaEnDiccionario(direccion) {
    const dic = cargarCallesZona()
    const nombres = extraerNombresCalles(direccion)
    for (const nombre of nombres) {
        // Búsqueda exacta
        if (dic[nombre]) return dic[nombre]
        // Búsqueda parcial: el nombre del diccionario está contenido en el de la dirección o viceversa
        for (const key of Object.keys(dic)) {
            if (key.length >= 4 && (nombre.includes(key) || key.includes(nombre))) {
                return dic[key]
            }
        }
        // Búsqueda fuzzy: similitud > 0.7 para cubrir errores ortográficos
        for (const key of Object.keys(dic)) {
            if (key.length >= 4 && similitudCalles(nombre, key) >= 0.7) {
                return dic[key]
            }
        }
    }
    return null
}

// Retorna { zona: 'norte'|'sur'|null, fuente: 'diccionario'|'geocodificacion'|'default' }
async function detectarZona(direccion, sock = null, contexto = '') {
    // 1) Diccionario de calles (rápido, sin red, ya cubre toda la ciudad)
    const zonaDic = buscarZonaEnDiccionario(direccion)
    if (zonaDic) {
        console.log(`📒 "${direccion}" -> zona ${zonaDic} (diccionario de calles)`)
        return { zona: zonaDic, fuente: 'diccionario', coords: null }
    }

    // 2) Fallback: geocodificación online (direcciones no encontradas en el diccionario)
    try {
        const coords = await geocodificar(direccion)
        if (coords) {
            const zona = detectarZonaPorCoordenadas(coords.lat, coords.lon)
            console.log(`📍 "${direccion}" -> ${coords.display} -> zona ${zona} (${coords.lat}, ${coords.lon})`)
            registrarCallesZona(direccion, zona)
            return { zona, fuente: 'geocodificacion', coords }
        }
    } catch (e) {
        console.log(`⚠️  Error geocodificando "${direccion}": ${e.message}`)
    }

    // No se pudo geocodificar → asignar zona del día automáticamente y aprender
    const zonaDelDiaActual = zonaDelDia(new Date()) || 'sur'
    console.log(`📅 "${direccion}" -> zona ${zonaDelDiaActual} (asignada por zona del día)`)

    // Guardar en diccionario para aprender
    registrarCallesZona(direccion, zonaDelDiaActual)

    // Notificar que se asignó por zona del día (no pedir revisión manual)
    if (sock) {
        try {
            const jidAlerta = `${NUMERO_PRUEBA}@s.whatsapp.net`
            await sock.sendMessage(jidAlerta, {
                text: `📅 *Dirección agregada al diccionario — zona ${zonaDelDiaActual.toUpperCase()}*\n\n📍 ${direccion}\n${contexto ? `📋 ${contexto}\n` : ''}\n_Se asignó la zona del día (${zonaDelDiaActual}). Si es incorrecta, corregir en calles_zona.json._`
            })
        } catch (e) {
            console.error('Error enviando notificación zona:', e.message)
        }
    }

    return { zona: zonaDelDiaActual, fuente: 'dia', coords: null }
}

// ===== CÁLCULO DE FECHAS =====
function esDiaHabil(fecha) {
    const dia = fecha.getDay()
    if (dia === 0) return false
    const fechaStr = fecha.toISOString().split('T')[0]
    return !FERIADOS.includes(fechaStr)
}

function proximoDiaCarga(zona) {
    // Norte carga: L(1), M(3), V(5)
    // Sur carga: M(2), J(4), S(6)
    const diasNorte = [1, 3, 5]
    const diasSur = [2, 4, 6]
    const diasZona = zona === 'norte' ? diasNorte : diasSur

    const ahora = new Date()

    // Si HOY es día de carga de esta zona y la orden de hoy todavía no se generó (antes de las 12:50),
    // el pedido entra en la orden de HOY.
    const hoy = new Date(ahora)
    hoy.setHours(0, 0, 0, 0)
    const horaActual = ahora.getHours()
    const minutoActual = ahora.getMinutes()
    const ordenDeHoyYaGenerada = (horaActual > 12) || (horaActual === 12 && minutoActual >= 50)

    if (diasZona.includes(hoy.getDay()) && esDiaHabil(hoy) && !ordenDeHoyYaGenerada) {
        return hoy
    }

    // Caso contrario, buscar el próximo día de carga (desde mañana)
    for (let i = 1; i <= 14; i++) {
        const fecha = new Date(ahora)
        fecha.setDate(ahora.getDate() + i)
        fecha.setHours(0, 0, 0, 0)
        if (diasZona.includes(fecha.getDay()) && esDiaHabil(fecha)) {
            return fecha
        }
    }
    return null
}

function calcularVencimiento(fechaCarga) {
    // 72hs = 3 días corridos desde las 16hs del día de carga
    // Salteando domingos y feriados
    let fecha = new Date(fechaCarga)
    fecha.setHours(16, 0, 0, 0)
    let diasContados = 0

    while (diasContados < 3) {
        fecha.setDate(fecha.getDate() + 1)
        const dia = fecha.getDay()
        const fechaStr = fecha.toISOString().split('T')[0]
        const esFeriado = FERIADOS.includes(fechaStr)
        const esDomingo = dia === 0
        if (!esFeriado && !esDomingo) diasContados++
    }
    fecha.setHours(16, 0, 0, 0)
    return fecha
}

function formatearFecha(fecha) {
    return new Date(fecha).toLocaleDateString('es-AR', {
        weekday: 'long', day: 'numeric', month: 'long',
        hour: '2-digit', minute: '2-digit'
    })
}

// ===== ANÁLISIS DE FOTOS CON OLLAMA =====
async function analizarFoto(buffer, tipoPedido) {
    try {
        const base64 = buffer.toString('base64')
        let pregunta = ''

        if (tipoPedido === 'Vehículo arrojando basura') {
            pregunta = `You are analyzing a photo for a municipal complaint about a vehicle illegally dumping waste. Look CAREFULLY at the image.

Step 1: Is there a vehicle visible in the image? (car, truck, motorcycle, etc.)
Step 2: If yes, is the license plate visible AND readable?

Be STRICT. If you cannot clearly see a vehicle and its license plate, mark as invalid.

Answer ONLY with valid JSON, no other text:
{"esValida": true or false, "razon": "explanation in Spanish, max 15 words"}`

        } else if (tipoPedido === 'Persona arrojando basura') {
            pregunta = `You are analyzing a photo for a municipal complaint about a person illegally dumping waste. Look CAREFULLY at the image.

Step 1: Is there a person clearly visible in the image?
Step 2: Can you see enough detail to potentially identify them or their location (clothing, surroundings, distinguishing features)?

Be STRICT. A photo showing only feet, floor, or unclear/blurry content with no identifiable person should be marked as invalid.

Answer ONLY with valid JSON, no other text:
{"esValida": true or false, "razon": "explanation in Spanish, max 15 words"}`

        } else {
            pregunta = `You are analyzing a photo for a municipal waste management complaint in Argentina (basural, residuos, bolsas, escombros, poda, muebles, etc).

Look at the image and determine if it shows:
- Garbage bags or waste of any kind
- Rubble, debris or construction waste
- Tree branches or green waste
- Furniture or bulky items
- Any accumulation of waste or environmental problem
- A street, sidewalk or public space with any waste

Be PERMISSIVE: if there is ANY doubt that this could be a waste-related photo, mark as valid.
Only mark as invalid if the photo clearly shows something completely unrelated (selfie, food, animals, etc).

Answer ONLY with valid JSON, no other text:
{"esValida": true or false, "razon": "explanation in Spanish, max 15 words"}`
        }

        const response = await ollama.chat({
            model: 'llava:13b',
            messages: [{ role: 'user', content: pregunta, images: [base64] }],
            options: { temperature: 0.1 }
        })

        const texto = response.message.content
        const match = texto.match(/\{[\s\S]*?\}/)
        if (match) return JSON.parse(match[0])
        return { esValida: false, razon: 'No se pudo verificar la foto, enviá otra' }
    } catch (e) {
        console.error('Error analizando foto:', e.message || e)
        // Si Ollama falla (timeout, error de red, etc.), aprobamos la foto
        // para no bloquear el pedido del vecino
        console.log('⚠️ Ollama no disponible — foto aprobada por defecto')
        return { esValida: true, razon: 'Foto recibida correctamente' }
    }
}

// ===== MENSAJES =====
const MSG = {
    bienvenida: `¡Hola! 👋 Bienvenido/a a la *Secretaría de Ambiente y Desarrollo Sustentable* de la Municipalidad de San Miguel de Tucumán.

_En cualquier momento escribí *menú* para volver al inicio._

Estamos acá para ayudarte. ¿Con qué podemos asistirte hoy?

1️⃣ Residuos no habituales
2️⃣ Basural o volcadero
3️⃣ Barrido
4️⃣ Falta de recolección
5️⃣ Persona o vehículo arrojando basura
6️⃣ Hablar con un asesor`,

    conversacionAnterior: (tipo) => `¡Hola! 👋 Veo que teníamos una conversación en curso sobre *${tipo}*.

¿Qué querés hacer?

1️⃣ Iniciar una consulta nueva
2️⃣ Continuar con lo que estábamos`,

    fueraDeHorario: `¡Hola! 👋 Gracias por comunicarte con la *Secretaría de Ambiente*.

En este momento estamos fuera del horario de atención con asesores.
🕗 Nuestro horario es *lunes a viernes de 8 a 13 hs*.

Tu mensaje quedó registrado y te respondemos en cuanto estemos disponibles 🙏`,

    opcion1: `¡Perfecto! Podemos ayudarte con el retiro de residuos no habituales 💪

Antes de continuar, ¿qué tipo de residuo necesitás retirar?

1️⃣ Madera / restos de poda / ramas
2️⃣ Escombros / material de construcción
3️⃣ Muebles / electrodomésticos / colchones
4️⃣ Otro tipo de residuo`,

    opcion1_madera: `Perfecto, podemos gestionar el retiro de madera y restos de poda 🌿

*Límites del servicio:*
• Restos de poda: hasta 10 bolsas, o ramas enfardadas hasta 1m³
• Si la cantidad supera estos límites, *no podremos cumplir el pedido*

Para continuar necesitamos:
📸 *Una foto clara* de lo que necesitás retirar
📍 *Tu dirección exacta*

¿Podés enviarnos esos datos?`,

    opcion1_escombros: `Podemos gestionar el retiro de escombros bajo las siguientes condiciones ⚠️

*Requisitos obligatorios:*
• Los escombros deben estar *embolsados* (hasta 5 bolsas de hasta 15kg cada una)
• Si la cantidad es mayor o no están embolsados, *no se puede cumplir el pedido*
• Para volúmenes grandes que requieren máquina, es un trámite separado que requiere tu *autorización escrita* ya que la vereda puede sufrir daños

Para continuar necesitamos:
📸 *Una foto clara* mostrando las bolsas de escombros
📍 *Tu dirección exacta*

¿Podés enviarnos esos datos?`,

    opcion1_muebles: `Podemos gestionar el retiro de muebles, electrodomésticos o colchones 🛋️

*Límites del servicio:*
• Hasta 1m³ de volumen total
• Si la cantidad supera este límite, *no podremos cumplir el pedido*

Para continuar necesitamos:
📸 *Una foto clara* de lo que necesitás retirar
📍 *Tu dirección exacta*

¿Podés enviarnos esos datos?`,

    opcion1_otro: `Podemos intentar gestionar el retiro 📋

*Tené en cuenta:*
• El volumen máximo que podemos retirar es de *1m³*
• Si supera ese límite o requiere equipamiento especial, *no podremos cumplir el pedido*

Para continuar necesitamos:
📸 *Una foto clara* de lo que necesitás retirar (es muy importante para evaluar el pedido)
📍 *Tu dirección exacta*

¿Podés enviarnos esos datos?`,

    opcion2: `Entendemos la situación y vamos a gestionarlo 🙏

Para poder actuar necesitamos:

📸 *Una foto clara y legible del lugar*
📍 *La dirección exacta*

¿Podés enviarnos esos datos?`,

    opcion3: `Tomamos nota 📝 Vamos a gestionar tu reclamo de barrido.

Para poder registrarlo necesitamos:

📸 *Una foto clara y legible de la zona afectada*
📍 *La dirección exacta*

¿Nos las podés enviar?`,

    opcion4: `Lamentamos el inconveniente, lo vamos a resolver 🙏

Para gestionar tu reclamo necesitamos:

📸 *Una foto clara y legible*
📍 *Tu dirección exacta*
🗓️ *¿Desde cuándo no pasa el servicio de recolección?*

También podés consultar los horarios según tu zona acá:
👉 https://smtendatos.gob.ar/mapa-interactivo-de-recoleccion-de-residuos-por-turno/`,

    opcion5: `Gracias por reportarlo 🙌

¿De qué se trata?

1️⃣ Vehículo arrojando basura
2️⃣ Persona arrojando basura`,

    opcion5a: `Para darle curso al reclamo necesitamos:

📸 *Foto del vehículo* — es importante que la *patente sea clara y legible*
📍 *Dirección exacta del hecho*
📅 *Fecha y hora en que ocurrió*

¿Podés enviarnos esos datos?`,

    opcion5b: `Para registrar el reclamo necesitamos:

📸 *Foto de la persona* — es importante que se pueda *identificar de dónde es*
📍 *Dirección exacta donde ocurrió*
📅 *Fecha y hora en que ocurrió*
👤 *Datos de la persona* (si los tenés)

¿Podés enviarnos esa información?`,

    fotoValida: `✅ *Foto verificada correctamente*

Ahora necesitamos tu dirección exacta para completar el pedido 📍`,

    fotoInvalida: (razon) => `⚠️ No pudimos verificar tu foto correctamente.

*${razon}*

Por favor enviá una foto más clara y legible para poder gestionar tu reclamo 📸`,

    fotoRecibidaSinDireccion: `✅ *Foto verificada correctamente*

Para completar tu pedido también necesitamos:

📍 *Tu dirección exacta*

Recordá que la foto tiene que ser *clara y legible* para poder gestionar correctamente tu reclamo.`,

    fotoRecibidaSinDireccionVehiculo: `✅ *Foto verificada correctamente*

Para completar tu reclamo también necesitamos:

📍 *Dirección exacta del hecho*
📅 *Fecha y hora en que ocurrió*

Recordá que en la foto debe verse la *patente del vehículo de forma clara y legible*.`,

    fotoRecibidaSinDireccionPersona: `✅ *Foto verificada correctamente*

Para completar tu reclamo también necesitamos:

📍 *Dirección exacta donde ocurrió*
📅 *Fecha y hora en que ocurrió*

Recordá que en la foto debe poder *identificarse de dónde es la persona*.`,

    pedidoActivo: `Tu pedido ya está registrado y en proceso 🙌

¿Qué necesitás?

1️⃣ Hacer un nuevo pedido
2️⃣ Ver mi pedido actual
3️⃣ Hablar con un asesor`,

    verPedido: (tipo, direccion, fechaOrden, fechaVence) =>
        `📋 *Tu pedido registrado:*\n\n` +
        `📌 Tipo: ${tipo}\n` +
        `📍 Dirección: ${direccion}\n` +
        `📅 Fecha de orden: ${formatearFecha(fechaOrden)}\n` +
        `⏱️ Plazo máximo: ${formatearFecha(fechaVence)}\n\n` +
        `Cuando el servicio se realice, avisanos escribiendo *"listo"* ✅`,

    preguntaCumplimiento: (fechaVence) =>
        `Hola 👋 El plazo de tu pedido venció el *${formatearFecha(fechaVence)}*.\n\n` +
        `¿El servicio fue realizado?\n\n` +
        `✅ Escribí *"sí"* si ya se realizó\n` +
        `❌ Escribí *"no"* si todavía no`,

    cumplido: `¡Muchas gracias por avisarnos! 🙏💚

Nos alegra que el servicio se haya realizado correctamente. Estamos siempre a tu disposición para lo que necesites.

¡Hasta pronto! 👋`,

    incumplido: `Lamentamos el inconveniente 🙏

Vamos a generar el reclamo pertinente para que la situación se solucione lo antes posible.

Un asesor va a tomar tu caso de inmediato.`,

    noCorresponde: `Gracias por comunicarte con la *Secretaría de Ambiente* 🙏

Tu consulta corresponde a la *SAT (Sociedad Aguas del Tucumán)*, que gestiona todo lo relacionado con agua, cloacas y servicios sanitarios.

💧 *SAT — Sociedad Aguas del Tucumán:*
🌐 https://www.sat.com.ar
📞 *0800-444-1726* (línea gratuita 24hs)

También podés gestionar otros pedidos municipales a través de:
📱 *App Ciudad Digital:* https://ciudaddigital.smt.gob.ar/#/registro
☎️ *Atención Ciudadana:* 381 223-0573`,

    noEntiendo: `Disculpá, no entendí tu mensaje 😊

Por favor elegí una de las opciones:

1️⃣ Residuos no habituales
2️⃣ Basural o volcadero
3️⃣ Barrido
4️⃣ Falta de recolección
5️⃣ Persona o vehículo arrojando basura
6️⃣ Hablar con un asesor`,

    opcion6: `¡Por supuesto! 😊 En un momento un asesor va a estar disponible para ayudarte.

Por favor dejanos tu consulta acá y te respondemos a la brevedad.

⏰ Nuestro horario de atención es *lunes a viernes de 8 a 13 hs*.`
}

// ===== PALABRAS QUE NO CORRESPONDEN =====
const palabrasNoCorresponden = [
    'cloaca','cloacal','cloacas','desagüe','desague',
    'caño','caños','plomería','plomeria','plomero',
    'destrancar','destrabar','pozo','séptico','septico',
    'agua corriente','pérdida de agua','perdida de agua',
    'inundación','inundacion','zanja','alcantarilla','sanitaria','tanque'
]

function noCorresponde(texto) {
    const lower = texto.toLowerCase()
    return palabrasNoCorresponden.some(p => lower.includes(p))
}

// ===== GESTIÓN DE PEDIDOS =====
function cargarPedidos() {
    if (fs.existsSync('pedidos.json')) return JSON.parse(fs.readFileSync('pedidos.json', 'utf8'))
    return []
}

function guardarPedidos(pedidos) {
    fs.writeFileSync('pedidos.json', JSON.stringify(pedidos, null, 2))
}

async function guardarPedido(numero, tipo, datos, sock = null) {
    const pedidos = cargarPedidos()
    const { zona, fuente: fuenteZona } = await detectarZona(datos.mensaje || '', sock, `${tipo} - ${numero}`)
    const fechaCarga = proximoDiaCarga(zona)
    const fechaVencimiento = calcularVencimiento(fechaCarga)

    const pedido = {
        id: Date.now(),
        numero, tipo, zona, zonaDetectadaPor: fuenteZona,
        fecha: new Date().toISOString(),
        fechaOrden: fechaCarga.toISOString(),
        fechaVencimiento: fechaVencimiento.toISOString(),
        estado: 'pendiente',
        fuente: 'Bot',
        avisoCumplimientoEnviado: false,
        ...datos
    }
    pedidos.push(pedido)
    guardarPedidos(pedidos)
    return pedido
}

function getPedidoActivo(numero) {
    const pedidos = cargarPedidos()
    return pedidos.find(p => p.numero === numero &&
        p.estado !== 'completado' && p.estado !== 'incumplido')
}

function cerrarPedido(numero, estadoCierre) {
    const pedidos = cargarPedidos()
    const idx = pedidos.findIndex(p => p.numero === numero &&
        (p.estado === 'pendiente' || p.estado === 'esperandoConfirmacion'))
    if (idx !== -1) {
        pedidos[idx].estado = estadoCierre
        pedidos[idx].fechaCierre = new Date().toISOString()
        guardarPedidos(pedidos)
    }
}

// ===== MAPEO LID -> NÚMERO REAL =====
// WhatsApp ya no expone el número real: entrega un ID interno (@lid). Cuando sí
// conocemos el número (vía remoteJidAlt, el evento lid-mapping.update, o
// resolviéndolo activamente con Baileys), lo persistimos acá para enriquecer
// el dashboard y la base electoral. Nunca se inventa un número.
function soloDigitos(jidOValor) {
    return (jidOValor || '').toString().replace(/@.*$/, '').replace(/\D/g, '')
}

function cargarLidMap() {
    if (fs.existsSync('lid_map.json')) {
        try { return JSON.parse(fs.readFileSync('lid_map.json', 'utf8')) } catch (e) { return {} }
    }
    return {}
}
function guardarLidMap(map) {
    fs.writeFileSync('lid_map.json', JSON.stringify(map, null, 2))
}
// Registra lid->pn (en dígitos). Devuelve true si hubo cambio.
function registrarLidMap(lidJid, pnJid) {
    const lid = soloDigitos(lidJid)
    const pn = soloDigitos(pnJid)
    if (!lid || !pn) return false
    const map = cargarLidMap()
    if (map[lid] === pn) return false
    map[lid] = pn
    guardarLidMap(map)
    return true
}
// Devuelve el número (dígitos) real de un @lid si lo conocemos, o null.
function buscarPNenMap(lidOJid) {
    const lid = soloDigitos(lidOJid)
    if (!lid) return null
    return cargarLidMap()[lid] || null
}
// Resolución activa vía Baileys (no bloquea: corre en background y persiste).
async function resolverPNdeLID(sock, lidJid) {
    try {
        if (!lidJid || !lidJid.includes('@lid')) return null
        const pn = await sock?.signalRepository?.lidMapping?.getPNForLID?.(lidJid)
        if (pn) { registrarLidMap(lidJid, pn); return soloDigitos(pn) }
    } catch (e) {}
    return null
}

// Da formato +54 9 a un número en dígitos (sin @). '' si no es válido.
function formatearNumeroLocal(digitos) {
    const limpio = (digitos || '').replace(/\D/g, '')
    if (!limpio) return ''
    const local = limpio.startsWith('549') ? limpio.slice(3) :
                  limpio.startsWith('54') ? limpio.slice(2) : limpio
    if (local.length === 10) {
        return `+54 9 ${local.slice(0,4)} ${local.slice(4,7)}-${local.slice(7)}`
    }
    return '+54 9 ' + local
}

// ===== CONVERSIÓN DE JID A NÚMERO LEGIBLE =====
function jidANumero(jid, jidAlt) {
    // 1) Si tenemos el JID alternativo (numero real, formato @s.whatsapp.net), usarlo
    let fuente = (jidAlt && jidAlt.includes('@s.whatsapp.net')) ? jidAlt : jid
    if (!fuente) return ''
    // 2) Si solo tenemos @lid, intentar resolver con el mapa persistido
    if (fuente.includes('@lid')) {
        const pn = buscarPNenMap(fuente)
        if (!pn) return '' // no conocemos el numero real, no se inventa
        fuente = pn // ya son dígitos
    }
    return formatearNumeroLocal(fuente.replace(/@.+$/, ''))
}

// ===== ESTADOS =====
let estados = {}

function cargarEstados() {
    if (fs.existsSync('estados.json')) estados = JSON.parse(fs.readFileSync('estados.json', 'utf8'))
}

function guardarEstados() {
    fs.writeFileSync('estados.json', JSON.stringify(estados, null, 2))
}

// ===== VERIFICAR VENCIMIENTOS =====
async function verificarVencimientos(sock) {
    const pedidos = cargarPedidos()
    const ahora = new Date()

    for (const pedido of pedidos) {
        if (pedido.estado !== 'pendiente') continue
        if (pedido.avisoCumplimientoEnviado) continue
        const vencimiento = new Date(pedido.fechaVencimiento)
        if (ahora >= vencimiento) {
            const jid = `${pedido.numero}@s.whatsapp.net`
            try {
                await sock.sendMessage(jid, { text: MSG.preguntaCumplimiento(vencimiento) })
                const idx = pedidos.findIndex(p => p.id === pedido.id)
                pedidos[idx].avisoCumplimientoEnviado = true
                pedidos[idx].estado = 'esperandoConfirmacion'
                guardarPedidos(pedidos)
                estados[jid] = { paso: 'esperandoConfirmacionCumplimiento' }
                guardarEstados()
            } catch (e) {
                console.error('Error enviando aviso:', e)
            }
        }
    }
}

// ===== CONFIRMAR PEDIDO =====
async function confirmarPedido(sock, jid, numero, tipo, datos) {
    if (esPersonalInterno(numero)) return // no generar pedidos de personal interno
    const pedido = await guardarPedido(numero, tipo, datos, sock)
    const msgConf =
        `¡Gracias! Recibimos tu solicitud completa 🙌\n\n` +
        `📋 Tu pedido ingresará en la orden del *${formatearFecha(pedido.fechaOrden)}*\n` +
        `⏱️ Plazo máximo: *${formatearFecha(pedido.fechaVencimiento)}*\n\n` +
        `Cuando el servicio se realice, escribinos *"listo"* ✅\n\n` +
        `Cualquier consulta, acá estamos 💚`
    await sock.sendMessage(jid, { text: msgConf })
    estados[jid] = { paso: 'inicio' }
    guardarEstados()
}

// ===== CLASIFICACIÓN MENSAJES BAM =====
function clasificarMensajeBAM(texto) {
    const lower = (texto || '').toLowerCase()
    if (lower.includes('volcadero')) return { categoria: 'volcadero', tipo: 'Volcadero' }
    if (lower.includes('basural')) return { categoria: 'basural', tipo: 'Basural' }
    if (lower.includes('cesto')) return { categoria: 'limpieza_cestos', tipo: 'Limpieza de cesto y basura dispersa' }
    if (lower.includes('rnh') || lower.includes('no habitual') || lower.includes('escombro') || lower.includes('poda') || lower.includes('colchon') || lower.includes('colchón')) return { categoria: 'rnh', tipo: 'RNH' }
    if (lower.includes('barrido')) return { categoria: 'barrido', tipo: 'Barrido de calles' }
    if (lower.includes('recolecc')) return { categoria: 'otros', tipo: 'Falta de recolección' }
    if (lower.includes('neumat') || lower.includes('neumát')) return { categoria: 'otros', tipo: 'RNH Retiro de neumáticos' }
    return { categoria: 'otros', tipo: null }
}

// Palabras de TIPO DE SERVICIO que aparecen al INICIO del mensaje
const TIPOS_INICIO = [
    'volcadero', 'basural', 'rnh', 'residuos no habituales?',
    'limpieza de cesto[s]?( y basura dispersa)?',
    'barrido manual', 'barrido( de calles)?',
    'falta de recolecci[oó]n', 'descacharreo',
    'recolecci[oó]n'
]

// Frases de DESCRIPCIÓN/ACCIÓN que aparecen al FINAL del mensaje, tras un punto
const DESCRIPCIONES_FINAL = [
    'levante de (restos de )?(poda|bolsas)', 'levante',
    'basural', 'volcadero', 'poda', 'escombro[s]?',
    'colch[oó]n', 'descacharreo', 'basura dispersa( en vereda)?',
    'rnh.*'
]

// Mensajes que NO son pedidos de limpieza (son para multas/informes)
function esMensajeDeInforme(texto) {
    if (/^\s*para informar\b/i.test(texto)) return true
    if (/dominio\s+[A-Z0-9]{5,8}/i.test(texto)) return true
    if (/informe\s+de\s+novedad/i.test(texto)) return true
    return false
}

// Extrae dirección/situación de un mensaje "Para informar..." (multa)
function extraerDatosMulta(texto) {
    let s = texto.replace(/[\n\r]+/g, ' ').trim()
    // Quitar el prefijo "Para informar"
    s = s.replace(/^\s*para informar\s*,?\s*/i, '').trim()

    // Buscar "por calle X" o "calle X" al final como ubicación
    const matchCalle = s.match(/(?:por\s+)?calle\s+([A-Za-záéíóúÁÉÍÓÚñÑ0-9\s°ª]+?)[.\s]*$/i)
    let ubicacion = matchCalle ? matchCalle[1].trim() : null

    return {
        situacion: s,
        ubicacion: ubicacion || 'Sin especificar'
    }
}

function extraerDireccionBAM(texto) {
    if (!texto) return null
    let s = texto.replace(/[\n\r]+/g, ' ').trim()

    if (esMensajeDeInforme(s)) return null

    // Formato con tabs: "Volcadero\twsp\tDirección" o "Tipo\tDirección"
    if (s.includes('\t')) {
        const partesTabs = s.split('\t').map(p => p.trim()).filter(Boolean)
        // Buscar la parte que parece una dirección (tiene número o "y" o "entre")
        const dirTab = partesTabs.find(p =>
            /\d/.test(p) || /\s(y|e|entre|esq)\s/i.test(p) ||
            /[Aa]v\.|[Pp]je\.|[Bb]arrio/i.test(p)
        )
        if (dirTab && !TIPOS_INICIO.some(t => new RegExp(`^${t}$`, 'i').test(dirTab)) && dirTab.toLowerCase() !== 'wsp') {
            return dirTab.replace(/[.,;:\s]+$/, '').trim()
        }
    }

    // Formato "Tipo. Dirección" (con punto después del tipo)
    for (const tipo of TIPOS_INICIO) {
        const re = new RegExp(`^\\s*${tipo}[.:]\\s*`, 'i')
        if (re.test(s)) {
            s = s.replace(re, '').trim()
            break
        }
    }

    // 1) Si empieza con un tipo de servicio (sin punto), quitarlo
    for (const tipo of TIPOS_INICIO) {
        const re = new RegExp(`^\\s*${tipo}\\s+`, 'i')
        if (re.test(s)) {
            s = s.replace(re, '').trim()
            break
        }
    }

    // 2) Si el texto tiene un punto y lo que sigue después es una descripción
    //    de servicio/acción (ej: "...hasta Pje. Lugones. Levante de bolsas"),
    //    cortar desde ese punto.
    //    OJO: "Pje." también tiene punto, por eso evaluamos desde el ÚLTIMO punto
    //    hacia atrás y verificamos si lo que sigue matchea una descripción final.
    const partes = s.split('.')
    if (partes.length > 1) {
        const ultimaParte = partes[partes.length - 1].trim()
        const esDescripcionFinal = DESCRIPCIONES_FINAL.some(d =>
            new RegExp(`^\\s*[a-z]*\\/?\\s*${d}\\s*$`, 'i').test(ultimaParte)
        )
        if (esDescripcionFinal && ultimaParte.length > 0) {
            s = partes.slice(0, -1).join('.').trim()
        }
    }

    // Limpiar puntuación residual al final
    s = s.replace(/[.,;:\s]+$/, '').trim()

    return s.length > 3 ? s : null
}

// ===== PROCESAR MENSAJE DEL GRUPO BAM =====
function cargarMultas() {
    if (fs.existsSync('multas.json')) return JSON.parse(fs.readFileSync('multas.json', 'utf8'))
    return []
}
function guardarMultas(multas) {
    fs.writeFileSync('multas.json', JSON.stringify(multas, null, 2))
}

// Desempaqueta mensajes envueltos en contenedores de WhatsApp:
// "mensajes temporales" (ephemeral), "ver una vez" (viewOnce), documentos con
// caption y mensajes editados. SIN esto, en un grupo con mensajes temporales
// activados NO se captura absolutamente nada.
function desempaquetarMensaje(message) {
    let m = message
    for (let i = 0; i < 5 && m; i++) {
        const inner =
            m.ephemeralMessage?.message ||
            m.viewOnceMessage?.message ||
            m.viewOnceMessageV2?.message ||
            m.viewOnceMessageV2Extension?.message ||
            m.documentWithCaptionMessage?.message ||
            m.editedMessage?.message ||
            null
        if (!inner) break
        m = inner
    }
    return m || {}
}

// Extrae el texto/caption de cualquier tipo de mensaje (ya desempaquetado)
function extraerTextoMensaje(contenido) {
    return (
        contenido?.imageMessage?.caption ||
        contenido?.videoMessage?.caption ||
        contenido?.documentMessage?.caption ||
        contenido?.conversation ||
        contenido?.extendedTextMessage?.text ||
        ''
    ).trim()
}

async function procesarMensajeBAM(sock, msg) {
    // Procesar tanto mensajes con foto (caption) como mensajes de texto puro,
    // desempaquetando primero cualquier envoltorio (ephemeral / viewOnce / etc.)
    const contenido = desempaquetarMensaje(msg.message)
    const tieneFoto = !!(contenido?.imageMessage || contenido?.documentMessage?.mimetype?.startsWith?.('image/'))
    const caption = extraerTextoMensaje(contenido)

    if (!caption) return

    // Dedupe idempotente: no reprocesar el mismo mensaje (notify duplicado o
    // backfill de historial). Clave estable = msg.key.id.
    const msgId = msg.key?.id
    if (msgId) {
        const yaCargados = cargarPedidos()
        if (yaCargados.some(p => p.msgId === msgId) || cargarMultas().some(m => m.msgId === msgId)) {
            return
        }
    }

    // Ignorar comandos del sistema enviados desde números de prueba
    if (/^generar orden/i.test(caption)) return
    if (/^generar reporte/i.test(caption)) return

    const fecha = msg.messageTimestamp
        ? new Date(Number(msg.messageTimestamp) * 1000).toISOString()
        : new Date().toISOString()

    // ===== CASO: mensaje "Para informar" -> es multa =====
    if (esMensajeDeInforme(caption)) {
        const { situacion, ubicacion } = extraerDatosMulta(caption)

        // Guardar foto en carpeta de multas solo si tiene imagen
        if (tieneFoto) {
            try {
                const carpetaMultas = path.join('dataset', 'multas')
                if (!fs.existsSync(carpetaMultas)) fs.mkdirSync(carpetaMultas, { recursive: true })
                const buffer = await downloadMediaMessage(msg, 'buffer', {})
                const nombreArchivo = `${fecha.replace(/[:.]/g, '-')}.jpg`
                fs.writeFileSync(path.join(carpetaMultas, nombreArchivo), buffer)
            } catch (e) {
                console.error('Error guardando foto multa:', e.message)
            }
        }

        const multas = cargarMultas()
        multas.push({
            id: Date.now(),
            msgId,
            fecha,
            mensaje: caption,
            situacion,
            ubicacion,
            tieneFoto,
            reporteGenerado: false
        })
        guardarMultas(multas)
        console.log(`🚔 [BAM-MULTA] ${ubicacion}: ${situacion.slice(0, 60)}`)
        return
    }

    // ===== CASO: pedido de limpieza normal =====
    const { categoria, tipo } = clasificarMensajeBAM(caption)
    if (!tipo) return

    const direccion = extraerDireccionBAM(caption)
    if (!direccion) return

    // Guardar foto en dataset solo si tiene imagen
    if (tieneFoto) {
        try {
            const carpetaCategoria = path.join('dataset', categoria)
            if (!fs.existsSync(carpetaCategoria)) fs.mkdirSync(carpetaCategoria, { recursive: true })
            const buffer = await downloadMediaMessage(msg, 'buffer', {})
            const nombreArchivo = `${fecha.replace(/[:.]/g, '-')}.jpg`
            fs.writeFileSync(path.join(carpetaCategoria, nombreArchivo), buffer)
        } catch (e) {
            console.error('Error guardando foto BAM:', e.message)
        }
    }

    const { zona, fuente: fuenteZona } = await detectarZona(direccion, sock, `BAM - ${caption.slice(0, 60)}`)

    const pedidos = cargarPedidos()
    pedidos.push({
        id: Date.now(),
        msgId,
        numero: 'BAM',
        tipo, zona, zonaDetectadaPor: fuenteZona,
        mensaje: caption,
        direccion,
        tieneFoto,
        fuente: 'BAM',
        fecha,
        estado: 'pendiente',
        avisoCumplimientoEnviado: false
    })
    guardarPedidos(pedidos)
    console.log(`📋 [BAM] [${tipo}/${zona}] ${direccion}`)
}

// ===== GENERADOR DE ORDEN DE SERVICIO =====
function formatearFechaHoy() {
    return new Date().toLocaleDateString('es-AR', { day: '2-digit', month: 'long', year: 'numeric' })
}

// Determina qué zona corresponde cargar en un día dado
// Norte carga: Lunes(1), Miércoles(3), Viernes(5)
// Sur carga: Martes(2), Jueves(4), Sábado(6)
// Domingos y feriados: no se carga
// Si el día siguiente hábil es feriado, se adelanta su zona
function zonaDelDia(fecha) {
    if (!esDiaHabil(fecha)) return null // domingo o feriado

    // Verificar si el próximo día hábil es feriado
    // En ese caso, este día reemplaza su zona por la del feriado
    const manana = new Date(fecha)
    manana.setDate(manana.getDate() + 1)

    // Buscar el próximo día hábil
    let proximoDiaHabil = new Date(manana)
    for (let i = 0; i < 7; i++) {
        const d = new Date(manana)
        d.setDate(manana.getDate() + i)
        const dStr = d.toISOString().split('T')[0]
        if (d.getDay() !== 0 && !FERIADOS.includes(dStr)) {
            // No es domingo ni feriado -> es hábil, no hay feriado inmediato
            break
        }
        if (d.getDay() !== 0 && FERIADOS.includes(dStr)) {
            // El día siguiente (o muy próximo) ES feriado
            // Retornar la zona de ese día feriado en lugar de la propia
            const diaFeriado = d.getDay()
            if ([1, 3, 5].includes(diaFeriado)) return 'norte'
            if ([2, 4, 6].includes(diaFeriado)) return 'sur'
            break
        }
    }

    // Zona normal del día
    const dia = fecha.getDay()
    if ([1, 3, 5].includes(dia)) return 'norte'
    if ([2, 4, 6].includes(dia)) return 'sur'
    return null
}

// Pedidos pendientes de la zona que corresponde cargar HOY,
// cuya fechaOrden sea hoy o anterior (no asignados a una orden todavía)
function pedidosParaOrden(pedidos, zona, fechaRef = new Date()) {
    const ref = new Date(fechaRef)
    ref.setHours(0, 0, 0, 0)

    return pedidos.filter(p => {
        if (p.estado !== 'pendiente') return false
        if (p.ordenGenerada) return false
        if (p.zona !== zona) return false

        if (p.fechaOrden) {
            const fechaOrden = new Date(p.fechaOrden)
            fechaOrden.setHours(0, 0, 0, 0)
            return fechaOrden <= ref
        }
        // Pedidos del BAM sin fechaOrden: siempre incluir si son de la zona
        return true
    })
}

// Lleva el correlativo de número de orden
function obtenerSiguienteNumeroOrden() {
    let contador = { ultimo: 134 }
    if (fs.existsSync('contador_ordenes.json')) {
        contador = JSON.parse(fs.readFileSync('contador_ordenes.json', 'utf8'))
    }
    contador.ultimo++
    fs.writeFileSync('contador_ordenes.json', JSON.stringify(contador, null, 2))
    return String(contador.ultimo).padStart(6, '0')
}

function limpiarDireccionOrden(texto) {
    if (!texto) return 'Sin especificar'
    return texto
        .replace(/@[0-9]+@?[a-z.]*/g, '') // quitar IDs @lid
        .replace(/\s+/g, ' ')
        .trim()
}

function formatearUbicacionOrden(pedido) {
    const direccion = limpiarDireccionOrden(pedido.direccion || pedido.mensaje)
    if (pedido.fuente === 'BAM') return `BAM, ${direccion}`
    if (pedido.fuente === 'AC') return `A.C. ${pedido.numeroAC || ''} ${direccion}`.trim()
    if (pedido.fuente === 'CMA') return `CMA, ${direccion}`
    // Ítems del Sheets con canal original visible
    if (pedido.canalSheet && pedido.canalSheet !== 'Wsp') return `${pedido.canalSheet}, ${direccion}`
    return `Wsp, ${direccion}`
}

function formatearDeficienciaOrden(pedido) {
    const tiposMap = {
        'RNH': 'RNH', 'Basural': 'Basural', 'Volcadero': 'Volcadero',
        'Barrido': 'Barrido de calles', 'Falta de recolección': 'Falta de recolección',
        'Vehículo arrojando basura': 'Persona/vehículo arrojando basura',
        'Persona arrojando basura': 'Persona arrojando basura',
        'Limpieza de cesto y basura dispersa': 'Limpieza de cesto y basura dispersa'
    }
    return tiposMap[pedido.tipo] || pedido.tipo
}

async function generarOrdenServicioZona(sock, zona, fecha) {
    const pedidos = cargarPedidos()
    const itemsBAM = pedidosParaOrden(pedidos, zona, fecha)

    // Leer ítems de Google Sheets para la fecha y zona
    let itemsSheets = []
    try {
        itemsSheets = await leerItemsSheets(fecha, zona)
    } catch(e) {
        console.error('Error leyendo Sheets:', e.message)
    }

    // Combinar: primero BAM, luego Sheets
    // Evitar duplicados por dirección similar
    const itemsSheetsNuevos = itemsSheets.filter(s => {
        const dirS = s.direccion.toLowerCase().trim()
        return !itemsBAM.some(b => {
            const dirB = (b.direccion || '').toLowerCase().trim()
            return dirB === dirS || (dirB.length > 5 && dirS.includes(dirB)) || (dirS.length > 5 && dirB.includes(dirS))
        })
    })

    const items = [...itemsBAM, ...itemsSheetsNuevos]

    if (items.length === 0) {
        console.log(`⚠️  [12:50] No hay pedidos pendientes de zona ${zona.toUpperCase()} para generar la orden de hoy.`)
        return
    }

    console.log(`📋 Orden ${zona.toUpperCase()}: ${itemsBAM.length} del BAM + ${itemsSheetsNuevos.length} del Sheets = ${items.length} total`)

    const fechaHoy = fecha.toLocaleDateString('es-AR', { day: '2-digit', month: 'long', year: 'numeric' })
    const numeroOrden = obtenerSiguienteNumeroOrden()
    const zonaLabel = zona === 'norte' ? 'ZONA NORTE' : 'ZONA SUR'
    const border = { style: BorderStyle.SINGLE, size: 1, color: "000000" }
    const borders = { top: border, bottom: border, left: border, right: border }
    const colWidths = [600, 3200, 4560, 500, 500]

    const headerRow1 = new TableRow({
        children: [
            new TableCell({ borders, width: { size: colWidths[0], type: WidthType.DXA },
                margins: { top: 80, bottom: 80, left: 100, right: 100 },
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Nº", bold: true })] })] }),
            new TableCell({ borders, width: { size: colWidths[1], type: WidthType.DXA },
                margins: { top: 80, bottom: 80, left: 100, right: 100 },
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Deficiencias detectadas: Servicio", bold: true })] })] }),
            new TableCell({ borders, width: { size: colWidths[2], type: WidthType.DXA },
                margins: { top: 80, bottom: 80, left: 100, right: 100 },
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Ubicación de la Falta del Servicio", bold: true })] })] }),
            new TableCell({ borders, width: { size: colWidths[3] + colWidths[4], type: WidthType.DXA }, columnSpan: 2,
                margins: { top: 80, bottom: 80, left: 100, right: 100 },
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Cumplió", bold: true })] })] }),
        ]
    })

    const headerRow2 = new TableRow({
        children: [
            new TableCell({ borders, width: { size: colWidths[0], type: WidthType.DXA },
                margins: { top: 40, bottom: 40, left: 100, right: 100 }, children: [new Paragraph("")] }),
            new TableCell({ borders, width: { size: colWidths[1], type: WidthType.DXA },
                margins: { top: 40, bottom: 40, left: 100, right: 100 }, children: [new Paragraph("")] }),
            new TableCell({ borders, width: { size: colWidths[2], type: WidthType.DXA },
                margins: { top: 40, bottom: 40, left: 100, right: 100 }, children: [new Paragraph("")] }),
            new TableCell({ borders, width: { size: colWidths[3], type: WidthType.DXA },
                margins: { top: 40, bottom: 40, left: 100, right: 100 },
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Si", bold: true })] })] }),
            new TableCell({ borders, width: { size: colWidths[4], type: WidthType.DXA },
                margins: { top: 40, bottom: 40, left: 100, right: 100 },
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "No", bold: true })] })] }),
        ]
    })

    // Función para crear filas de datos
    function crearFilas(itemsHoja, offsetIdx) {
        return itemsHoja.map((pedido, idx) => {
            const numero = String(offsetIdx + idx + 1).padStart(2, '0')
            return new TableRow({
                children: [
                    new TableCell({ borders, width: { size: colWidths[0], type: WidthType.DXA },
                        margins: { top: 60, bottom: 60, left: 100, right: 100 },
                        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun(numero)] })] }),
                    new TableCell({ borders, width: { size: colWidths[1], type: WidthType.DXA },
                        margins: { top: 60, bottom: 60, left: 100, right: 100 },
                        children: [new Paragraph({ children: [new TextRun(formatearDeficienciaOrden(pedido))] })] }),
                    new TableCell({ borders, width: { size: colWidths[2], type: WidthType.DXA },
                        margins: { top: 60, bottom: 60, left: 100, right: 100 },
                        children: [new Paragraph({ children: [new TextRun(formatearUbicacionOrden(pedido))] })] }),
                    new TableCell({ borders, width: { size: colWidths[3], type: WidthType.DXA },
                        margins: { top: 60, bottom: 60, left: 100, right: 100 }, children: [new Paragraph("")] }),
                    new TableCell({ borders, width: { size: colWidths[4], type: WidthType.DXA },
                        margins: { top: 60, bottom: 60, left: 100, right: 100 }, children: [new Paragraph("")] }),
                ]
            })
        })
    }

    // Función para crear encabezado de hoja
    function crearEncabezadoHoja(nroHoja, totalHojas) {
        const hojaLabel = totalHojas > 1 ? ` (hoja ${nroHoja})` : ''
        return [
            new Paragraph({ children: [new TextRun({ text: "MUNICIPALIDAD DE SAN MIGUEL DE TUCUMÁN", bold: true, size: 24 })] }),
            new Paragraph({ children: [new TextRun({ text: "9 de Julio 570 S.M.T", size: 22 })] }),
            new Paragraph({ children: [new TextRun("")] }),
            new Paragraph({ alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: "ORDEN DE SERVICIO INTERNA A LA EMPRESA: TRANSPORTES 9 DE JULIO S.A.", bold: true, size: 24 })] }),
            new Paragraph({ children: [new TextRun("")] }),
            new Paragraph({ alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: zonaLabel, bold: true, size: 24 })] }),
            new Paragraph({ children: [new TextRun("")] }),
            new Paragraph({ alignment: AlignmentType.RIGHT,
                children: [new TextRun({ text: `Nº: ${numeroOrden}${hojaLabel}`, bold: true, size: 24 })] }),
            new Paragraph({ children: [new TextRun("")] }),
            new Paragraph({ alignment: AlignmentType.RIGHT,
                children: [new TextRun({ text: `San Miguel de Tucumán, ${fechaHoy}.`, size: 22 })] }),
            new Paragraph({ children: [new TextRun("")] }),
            new Paragraph({ children: [new TextRun({
                text: "POR LA PRESENTE SE COMUNICA A UDS. QUE, EN EL PLAZO DE 24 HORAS DE RECIBIDA LA PRESENTE, DEBERÁN DAR SOLUCIÓN A LAS DEFICIENCIAS DETECTADAS EN EL DÍA DE LA FECHA, EN LAS ZONAS DE INSPECCIÓN Y LUGARES QUE SE DETALLAN A CONTINUACIÓN:",
                bold: true, size: 22
            })] }),
            new Paragraph({ children: [new TextRun("")] }),
        ]
    }

    // Dividir items en hojas de 31
    const ITEMS_POR_HOJA = 31
    const hojas = []
    for (let i = 0; i < items.length; i += ITEMS_POR_HOJA) {
        hojas.push(items.slice(i, i + ITEMS_POR_HOJA))
    }
    const totalHojas = hojas.length

    // Crear secciones para cada hoja
    const secciones = hojas.map((itemsHoja, hojaIdx) => {
        const offsetIdx = hojaIdx * ITEMS_POR_HOJA
        const dataRows = crearFilas(itemsHoja, offsetIdx)
        return {
            properties: {
                page: {
                    size: { width: 12240, height: 15840 },
                    margin: { top: 1080, right: 1440, bottom: 1080, left: 1440 }
                }
            },
            children: [
                ...crearEncabezadoHoja(hojaIdx + 1, totalHojas),
                new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: colWidths, rows: [headerRow1, headerRow2, ...dataRows] }),
                new Paragraph({ children: [new TextRun("")] }),
                new Paragraph({ children: [
                    new TextRun({ text: "Marcar en casilleros: ", bold: true }),
                    new TextRun({ text: "Si (    )  No (    )", bold: true })
                ] }),
            ]
        }
    })

    const doc = new Document({
        styles: { default: { document: { run: { font: "Arial", size: 22 } } } },
        sections: secciones
    })

    const buffer = await Packer.toBuffer(doc)
    const fechaArchivo = fecha.toISOString().split('T')[0]
    const nombreArchivo = `OrdenServ_${fechaArchivo}_${zona}.docx`
    fs.writeFileSync(nombreArchivo, buffer)

    const pedidosActualizados = pedidos.map(p => {
        if (items.includes(p)) {
            return { ...p, estado: 'en_proceso', ordenGenerada: true, numeroOrden, fechaOrdenGenerada: new Date().toISOString() }
        }
        return p
    })
    guardarPedidos(pedidosActualizados)

    console.log(`✅ Orden de servicio generada: ${nombreArchivo}`)
    console.log(`   📋 ${zonaLabel} - ${items.length} items incluidos (fuentes: ${[...new Set(items.map(i => i.fuente || 'Bot'))].join(', ')})`)

    // Enviar por WhatsApp
    if (sock && NUMERO_ORDEN_SERVICIO) {
        try {
            const jidDestino = `${NUMERO_ORDEN_SERVICIO}@s.whatsapp.net`
            await sock.sendMessage(jidDestino, {
                document: fs.readFileSync(nombreArchivo),
                mimetype: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                fileName: nombreArchivo,
                caption: `📋 Orden de Servicio - ${zonaLabel} - ${fechaHoy}\n${items.length} items (${[...new Set(items.map(i => i.fuente || 'Bot'))].join(', ')})`
            })
            console.log(`📤 Orden enviada por WhatsApp a +${NUMERO_ORDEN_SERVICIO}`)
        } catch (e) {
            console.error('Error enviando orden por WhatsApp:', e.message)
        }
    }
}

async function generarOrdenServicio(sock) {
    const hoy = new Date()
    const zona = zonaDelDia(hoy)
    if (!zona) {
        const esFeriado = FERIADOS.includes(hoy.toISOString().split('T')[0])
        const motivo = esFeriado ? 'es feriado' : 'es domingo'
        console.log(`⚠️  Hoy ${motivo}, no se carga ninguna zona. No se genera orden.`)
        if (sock) {
            try {
                await sock.sendMessage(`${NUMERO_PRUEBA}@s.whatsapp.net`, {
                    text: `⚠️ Hoy ${motivo}, no se genera orden de servicio.`
                })
            } catch(e) {}
        }
        return
    }
    await generarOrdenServicioZona(sock, zona, hoy)
}

// ===== GENERADOR DE REPORTE DE MULTAS =====
async function generarReporteMultas(sock) {
    const multas = cargarMultas()
    const hoy = new Date()
    const hoyStr = hoy.toISOString().split('T')[0]

    const items = multas.filter(m => {
        if (m.reporteGenerado) return false
        const fechaMulta = new Date(m.fecha).toISOString().split('T')[0]
        return fechaMulta === hoyStr
    })

    if (items.length === 0) {
        console.log('⚠️  [20:00] No hay reportes "para informar/multar" pendientes para hoy.')
        return
    }

    const fechaHoy = formatearFechaHoy()
    const border = { style: BorderStyle.SINGLE, size: 1, color: "000000" }
    const borders = { top: border, bottom: border, left: border, right: border }
    const colWidths = [600, 2500, 6260]

    const headerRow = new TableRow({
        children: [
            new TableCell({ borders, width: { size: colWidths[0], type: WidthType.DXA },
                margins: { top: 80, bottom: 80, left: 100, right: 100 },
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Nº", bold: true })] })] }),
            new TableCell({ borders, width: { size: colWidths[1], type: WidthType.DXA },
                margins: { top: 80, bottom: 80, left: 100, right: 100 },
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Ubicación", bold: true })] })] }),
            new TableCell({ borders, width: { size: colWidths[2], type: WidthType.DXA },
                margins: { top: 80, bottom: 80, left: 100, right: 100 },
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Situación reportada", bold: true })] })] }),
        ]
    })

    const dataRows = items.map((item, idx) => {
        const numero = String(idx + 1).padStart(2, '0')
        const hora = new Date(item.fecha).toLocaleTimeString('es-AR', { hour: '2-digit', minute: '2-digit' })
        return new TableRow({
            children: [
                new TableCell({ borders, width: { size: colWidths[0], type: WidthType.DXA },
                    margins: { top: 60, bottom: 60, left: 100, right: 100 },
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun(numero)] })] }),
                new TableCell({ borders, width: { size: colWidths[1], type: WidthType.DXA },
                    margins: { top: 60, bottom: 60, left: 100, right: 100 },
                    children: [new Paragraph({ children: [new TextRun(item.ubicacion)] })] }),
                new TableCell({ borders, width: { size: colWidths[2], type: WidthType.DXA },
                    margins: { top: 60, bottom: 60, left: 100, right: 100 },
                    children: [new Paragraph({ children: [new TextRun(`[${hora}hs] ${item.situacion}`)] })] }),
            ]
        })
    })

    const doc = new Document({
        styles: { default: { document: { run: { font: "Arial", size: 22 } } } },
        sections: [{
            properties: {
                page: {
                    size: { width: 12240, height: 15840 },
                    margin: { top: 1080, right: 1440, bottom: 1080, left: 1440 }
                }
            },
            children: [
                new Paragraph({ children: [new TextRun({ text: "MUNICIPALIDAD DE SAN MIGUEL DE TUCUMÁN", bold: true, size: 24 })] }),
                new Paragraph({ children: [new TextRun({ text: "Secretaría de Ambiente y Desarrollo Sustentable", size: 22 })] }),
                new Paragraph({ children: [new TextRun("")] }),
                new Paragraph({ alignment: AlignmentType.CENTER,
                    children: [new TextRun({ text: "REPORTE DIARIO DE SITUACIONES PARA INFORMAR / MULTAR", bold: true, size: 24 })] }),
                new Paragraph({ children: [new TextRun("")] }),
                new Paragraph({ alignment: AlignmentType.RIGHT,
                    children: [new TextRun({ text: `San Miguel de Tucumán, ${fechaHoy}.`, size: 22 })] }),
                new Paragraph({ children: [new TextRun("")] }),
                new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: colWidths, rows: [headerRow, ...dataRows] }),
            ]
        }]
    })

    const buffer = await Packer.toBuffer(doc)
    const fechaArchivo = new Date().toISOString().split('T')[0]
    const nombreArchivo = `Multas_${fechaArchivo}.docx`
    fs.writeFileSync(nombreArchivo, buffer)

    const multasActualizadas = multas.map(m => {
        if (items.includes(m)) return { ...m, reporteGenerado: true, fechaReporteGenerado: new Date().toISOString() }
        return m
    })
    guardarMultas(multasActualizadas)

    console.log(`✅ [20:00] Reporte de multas generado: ${nombreArchivo}`)
    console.log(`   🚔 ${items.length} situaciones incluidas`)

    if (sock && NUMERO_MULTAS) {
        try {
            const jidDestino = `${NUMERO_MULTAS}@s.whatsapp.net`
            await sock.sendMessage(jidDestino, {
                document: fs.readFileSync(nombreArchivo),
                mimetype: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                fileName: nombreArchivo,
                caption: `🚔 Reporte de Multas - ${fechaHoy}\n${items.length} situaciones reportadas`
            })
            console.log(`📤 Reporte de multas enviado por WhatsApp a +${NUMERO_MULTAS}`)
        } catch (e) {
            console.error('Error enviando reporte de multas:', e.message)
        }
    }
}

// ===== API DE RECLAMOS MUNICIPALES =====
const ID_PERSONA = 32015
const API_RECLAMOS = 'https://estadisticas.smt.gob.ar:8084/reclamos/traerReclamos'

function cargarReclamos() {
    if (fs.existsSync('reclamos.json')) return JSON.parse(fs.readFileSync('reclamos.json', 'utf8'))
    return []
}

function guardarReclamos(data) {
    fs.writeFileSync('reclamos.json', JSON.stringify(data, null, 2))
}

function consultarAPIReclamos() {
    return new Promise((resolve, reject) => {
        const hoy = new Date()
        const desde = new Date(hoy)
        desde.setMonth(desde.getMonth() - 6) // últimos 6 meses para visualización
        const desdeStr = desde.toISOString().split('T')[0]
        const hastaStr = hoy.toISOString().split('T')[0]

        const body = JSON.stringify({ desde: desdeStr, hasta: hastaStr, id_persona: ID_PERSONA })
        const options = {
            hostname: 'estadisticas.smt.gob.ar',
            port: 8084,
            path: '/reclamos/traerReclamos',
            method: 'POST',
            rejectUnauthorized: false, // certificado municipal autofirmado
            headers: {
                'Content-Type': 'application/json',
                'Content-Length': Buffer.byteLength(body)
            }
        }

        const req = https.request(options, (res) => {
            let data = ''
            res.on('data', c => data += c)
            res.on('end', () => {
                try {
                    resolve(JSON.parse(data))
                } catch(e) { reject(e) }
            })
        })
        req.on('error', reject)
        req.write(body)
        req.end()
    })
}

async function actualizarReclamos() {
    try {
        console.log('🔄 Consultando API de reclamos municipales...')
        const reclamos = await consultarAPIReclamos()

        if (!Array.isArray(reclamos)) {
            console.log('⚠️  Respuesta inesperada de la API de reclamos')
            return
        }

        guardarReclamos(reclamos)
        console.log(`✅ ${reclamos.length} reclamos actualizados desde la API municipal`)

        // Actualizar dashboard con datos cruzados
        actualizarDashboardData()

    } catch(e) {
        console.error('Error consultando API de reclamos:', e.message)
    }
}

// ===== GENERADOR DE DASHBOARD_DATA.JSON =====
function actualizarDashboardData() {
    try {
        const satisfaccion = cargarSatisfaccion()
        const pedidos = cargarPedidos()
        const reclamos = cargarReclamos()

        // Estadísticas de satisfacción
        const pos = satisfaccion.filter(s => s.sentimiento === 'positivo').length
        const neg = satisfaccion.filter(s => s.sentimiento === 'negativo').length
        const neu = satisfaccion.filter(s => s.sentimiento === 'neutro').length

        // Estadísticas de pedidos del día
        const hoy = new Date().toISOString().split('T')[0]
        const pedidosHoy = pedidos.filter(p => {
            const fecha = p.fechaOrden ? new Date(p.fechaOrden).toISOString().split('T')[0] : null
            return fecha === hoy || (!p.ordenGenerada && p.estado === 'pendiente')
        })
        const pedidosBot = pedidosHoy.filter(p => p.fuente === 'Bot').length
        const pedidosBAM = pedidosHoy.filter(p => p.fuente === 'BAM').length
        const completados = pedidos.filter(p => p.estado === 'completado').length
        const pendientes = pedidos.filter(p => p.estado === 'pendiente').length

        // Reclamos por estado (estados reales del sistema SMT)
        const reclamosIniciados = reclamos.filter(r => r.nombre_estado === 'INICIADO').length
        const reclamosProceso = reclamos.filter(r => r.nombre_estado === 'EN PROCESO').length
        const reclamosFinalizados = reclamos.filter(r =>
            r.nombre_estado === 'FINALIZADO' ||
            r.nombre_estado === 'FINALIZADO CON DERIVACION' ||
            r.nombre_estado === 'FINALIZADO CON DERIVACION EXT' ||
            r.nombre_estado === 'FINALIZADO SIN SOLUCION'
        ).length
        const reclamosDerivaados = reclamos.filter(r => r.nombre_estado === 'DERIVADO').length

        // Vecinos con datos completos (para base electoral)
        // Incluir vecinos del WhatsApp. El número real se resuelve con el mapa
        // LID->PN (pendiente 3); NUNCA se expone el @lid como teléfono.
        const vecinosWhatsApp = satisfaccion.map(s => {
            const pnReal = soloDigitos(s.celular) || buscarPNenMap(s.numero) || ''
            const ultimos8 = pnReal.length >= 8 ? pnReal.slice(-8) : null
            const reclamoCruzado = ultimos8 ? reclamos.find(r =>
                r.telefono && r.telefono.replace(/[^0-9]/g,'').slice(-8) === ultimos8
            ) : null
            const telefonoReal = reclamoCruzado?.telefono || (pnReal ? formatearNumeroLocal(pnReal) : '')
            return {
                ...s,
                canal: 'WhatsApp Ambiente',
                nombre: reclamoCruzado?.apellido_nombre || s.nombre || '',
                telefono: telefonoReal, // '' si no conocemos el número real (no se inventa)
                celular: telefonoReal || s.celular || '',
                email: reclamoCruzado?.email || '',
                direccion: reclamoCruzado?.direccion || s.direccion || '',
                lat: reclamoCruzado?.coorde1 || null,
                lon: reclamoCruzado?.coorde2 || null,
                distrito: reclamoCruzado?.DISTRITO || null
            }
        })

        // Incluir vecinos de la API que NO están ya en WhatsApp.
        // Dedup por número real resuelto (no por @lid).
        const telefonosWhatsApp = new Set(
            satisfaccion
                .map(s => soloDigitos(s.celular) || buscarPNenMap(s.numero) || '')
                .filter(t => t.length >= 8)
                .map(t => t.slice(-8))
        )
        const vecinosAPI = reclamos
            .filter(r => r.telefono && !telefonosWhatsApp.has(r.telefono.replace(/[^0-9]/g,'').slice(-8)))
            .map(r => ({
                numero: r.telefono,
                canal: r.nombre_oreclamo || 'App Ciudad Digital',
                nombre: r.apellido_nombre || '',
                telefono: r.telefono,
                email: r.email || '',
                direccion: r.direccion || '',
                lat: r.coorde1 || null,
                lon: r.coorde2 || null,
                distrito: r.DISTRITO || null,
                sentimiento: 'neutro', // sin análisis de sentimiento para reclamos API
                asunto: r.nombre_treclamo || '',
                area: r.nombre_oficina || '',
                estado: r.nombre_estado || '',
                fecha: r.fecha_hora_inicio || ''
            }))

        const vecinosCompletos = [...vecinosWhatsApp, ...vecinosAPI]

        // Canales
        const porCanal = {}
        vecinosCompletos.forEach(v => {
            const canal = v.canal || 'Otro'
            porCanal[canal] = (porCanal[canal] || 0) + 1
        })

        const dashData = {
            satisfaccion: vecinosCompletos,
            pedidos: {
                bot: pedidosBot,
                bam: pedidosBAM,
                completados,
                pendientes
            },
            reclamos: {
                total: reclamos.length,
                iniciados: reclamosIniciados,
                en_proceso: reclamosProceso,
                finalizados: reclamosFinalizados,
                derivados: reclamosDerivaados
            },
            canales: porCanal,
            totalVecinos: vecinosCompletos.length,
            actualizado: new Date().toISOString()
        }

        fs.writeFileSync('dashboard_data.json', JSON.stringify(dashData, null, 2))
        console.log('📊 dashboard_data.json actualizado')

        // Push automático a GitHub
        subirDashboardAGitHub()

    } catch(e) {
        console.error('Error actualizando dashboard:', e.message)
    }
}

// ===== LECTURA DE GOOGLE SHEETS =====
// Lee los ítems del Google Sheet desde sheets_cache.json.
// El cache lo refresca un proceso externo (Python) y guarda el contenido en
// formato markdown pipe-delimited, que es justo lo que parsearSheetMarkdown
// espera. El export CSV en vivo NO sirve: viene separado por comas y el parser
// nunca matchea las filas "| Orden de servicio ... |".
async function leerItemsSheets(fecha, zona) {
    try {
        if (!fs.existsSync('sheets_cache.json')) {
            console.warn('⚠️  sheets_cache.json no existe — no se incorporan ítems del Sheets')
            return []
        }
        const cache = JSON.parse(fs.readFileSync('sheets_cache.json', 'utf8'))
        const contenido = cache.contenido || ''
        if (!contenido.trim()) {
            console.warn('⚠️  sheets_cache.json vacío — no se incorporan ítems del Sheets')
            return []
        }

        // Avisar si el cache quedó viejo (proceso externo caído)
        if (cache.actualizado) {
            const horas = (Date.now() - new Date(cache.actualizado).getTime()) / 3.6e6
            if (horas > 24) {
                console.warn(`⚠️  sheets_cache.json desactualizado (${horas.toFixed(0)}h). Revisar el proceso que lo refresca.`)
            }
        }

        return parsearSheetMarkdown(contenido, fecha, zona)
    } catch (e) {
        console.error('Error leyendo sheets_cache.json:', e.message)
        return []
    }
}

function parsearSheetMarkdown(md, fecha, zona) {
    const items = []
    const lineas = md.split('\n').map(l => l.trim()).filter(Boolean)

    const d = fecha.getDate()
    const m = fecha.getMonth() + 1
    const fechaFormatos = [
        `${d}/${m}`,
        `${String(d).padStart(2,'0')}/${m}`,
        `${d}/${String(m).padStart(2,'0')}`,
        `${String(d).padStart(2,'0')}/${String(m).padStart(2,'0')}`
    ]
    const zonaLabel = zona.toLowerCase()

    // Parsear celda de tabla markdown: | contenido |
    function parseCeldas(linea) {
        return linea.split('|').map(c => c.trim()).filter((c, i, arr) => i > 0 && i < arr.length - 1)
    }

    let dentroDeOrden = false
    let saltarCabecera = false

    for (const linea of lineas) {
        // Ignorar separadores de tabla (:---:)
        if (/^\|[\s:\-|]+\|$/.test(linea)) continue

        const cols = parseCeldas(linea)
        if (!cols.length) continue
        const col0 = cols[0]

        // Detectar inicio de orden: | Orden de servicio NNN | fecha | | ZONA |
        if (/^Orden de servicio\s+\d+/i.test(col0)) {
            const fechaSheet = (cols[1] || '').trim()
            const zonaSheet = linea.toLowerCase()
            const esNorte = zonaSheet.includes('norte')
            const esSur = zonaSheet.includes('sur')
            const esZonaCorrecta = (zonaLabel === 'norte' && esNorte) || (zonaLabel === 'sur' && esSur)
            const esFechaCorrecta = fechaFormatos.includes(fechaSheet)
            dentroDeOrden = esFechaCorrecta && esZonaCorrecta
            saltarCabecera = dentroDeOrden
            continue
        }

        if (!dentroDeOrden) continue

        // Saltar cabecera (Asunto | | Dirección | Contacto)
        if (/^asunto$/i.test(col0)) { saltarCabecera = false; continue }
        if (saltarCabecera) { saltarCabecera = false; continue }

        const asunto = col0
        const canal = (cols[1] || '').trim()
        const direccion = (cols[2] || '').trim()
        const contacto = (cols[3] || '').trim()

        if (!asunto || !direccion) continue

        let tipo = 'RNH'
        const a = asunto.toLowerCase()
        if (/basural|basura en (la|v[ií]a)/.test(a)) tipo = 'Basural'
        else if (/volcadero/.test(a)) tipo = 'Volcadero'
        else if (/barrido/.test(a)) tipo = 'Barrido'
        else if (/falta de recolec/.test(a)) tipo = 'Falta de recolección'
        else if (/limpieza de cesto/.test(a)) tipo = 'Limpieza de cesto y basura dispersa'
        else if (/raspado|levante de barro/.test(a)) tipo = 'Raspado y levante de barro'
        else if (/levante de poda|levante de bolsas/.test(a)) tipo = 'Levante de poda'

        let fuenteLabel = 'Wsp'
        if (/^\d{5,}$/.test(canal)) fuenteLabel = `AC ${canal}`
        else if (/^ac\s/i.test(canal)) fuenteLabel = canal.toUpperCase()
        else if (/^cma$/i.test(canal)) fuenteLabel = 'CMA'

        items.push({
            tipo,
            direccion,
            fuente: fuenteLabel.startsWith('AC') ? 'AC' : (fuenteLabel === 'CMA' ? 'CMA' : 'Wsp'),
            numeroAC: fuenteLabel.startsWith('AC') ? canal.replace(/^ac\s*/i, '').trim() : '',
            canalSheet: fuenteLabel,
            contacto,
            asunto
        })
    }

    console.log(`📊 Google Sheets: ${items.length} ítems para ${fechaFormatos[0]} ${zona}`)
    return items
}

// ===== PUSH AUTOMÁTICO A GITHUB =====
const { exec } = require('child_process')

let pushEnCurso = false
function subirDashboardAGitHub() {
    if (pushEnCurso) return
    pushEnCurso = true
    exec('git add dashboard_data.json && git commit -m "actualizacion automatica" && git push origin master --force', 
        { cwd: 'C:\\wextractor' },
        (error, stdout, stderr) => {
            pushEnCurso = false
            if (error) {
                // Si no hay cambios es normal, no es error
                if (stderr && stderr.includes('nothing to commit')) {
                    return
                }
                console.error('Error push GitHub:', error.message)
                return
            }
            console.log('📤 Dashboard subido a GitHub')
        }
    )
}

// ===== SCHEDULER 12:50hs y 20:00hs =====
function programarGeneracionOrden(sock) {
    setInterval(() => {
        const ahora = new Date()
        if (ahora.getHours() === 12 && ahora.getMinutes() === 50) {
            generarOrdenServicio(sock)
        }
        if (ahora.getHours() === 20 && ahora.getMinutes() === 0) {
            generarReporteMultas(sock)
        }
        // Actualizar reclamos y dashboard cada 30 minutos
        if (ahora.getMinutes() === 0 || ahora.getMinutes() === 30) {
            actualizarReclamos().then(() => actualizarDashboardData())
        }
    }, 60 * 1000)
}

// ===== DETECCIÓN DE CUMPLIMIENTO DEL SERVICIO =====
function detectaCumplimiento(texto) {
    const t = texto.toLowerCase()
        .normalize('NFD').replace(/[̀-ͯ]/g, '') // quitar acentos
    
    const patrones = [
        // Confirmaciones directas
        /\blisto\b/, /\bok\b/, /\bconfirmo\b/, /\bconfirmado\b/,
        
        // Ya + acción de servicio
        /ya (paso|pasaron|vinieron|vino|levantaron|levanto|limpiaron|limpio|retiraron|retiro|sacaron|saco|lo hicieron|lo hizo|lo realizaron|realizo|cumplieron|cumplio|soluciono|solucionaron)/,
        
        // Servicio + completado
        /(el servicio|el trabajo|el retiro|la limpieza).*(se hizo|se realizo|se completo|fue realizado|fue hecho|lo hicieron|quedo listo)/,
        
        // Pasaron / vinieron sin "ya"
        /(pasaron|vinieron|vino|llegaron|llego).*(y (lo|la|los|las) (llevaron|sacaron|retiraron|limpiaron|levantaron))/,
        
        // Quedo + estado positivo
        /quedo (limpio|listo|solucionado|todo bien|impecable)/,
        
        // Se llevaron / retiraron
        /(se lo llevaron|se los llevaron|lo retiraron|los retiraron|lo sacaron|los sacaron|se lo sacaron)/,
        
        // Gracias + contexto de cumplimiento
        /(muchas gracias|muy amable|excelente).*(ya|cumplieron|lo hicieron|retiraron|pasaron|vinieron)/,
        
        // Todo + resultado positivo
        /(todo resuelto|todo solucionado|todo bien|quedo todo|se soluciono todo)/,
        
        // Variantes coloquiales tucumanas
        /(ya le dieron|ya lo dieron|ya lo hicieron|ya paso el camion|ya paso la empresa|ya paso el municipio)/
    ]
    
    return patrones.some(p => p.test(t))
}

// ===== ANÁLISIS DE SATISFACCIÓN (uso electoral) =====
function cargarSatisfaccion() {
    if (fs.existsSync('satisfaccion.json')) return JSON.parse(fs.readFileSync('satisfaccion.json', 'utf8'))
    return []
}
function guardarSatisfaccion(data) {
    fs.writeFileSync('satisfaccion.json', JSON.stringify(data, null, 2))
}

async function analizarSentimiento(texto) {
    try {
        const pregunta = `Analiza el siguiente mensaje que un vecino le escribió a la Secretaría de Ambiente Municipal sobre el servicio de recolección de residuos.

Clasificalo en UNA de estas 3 categorías:
- "positivo": agradecimiento, elogio, satisfacción con el servicio (ej: "gracias", "excelente", "muy rápido", "qué bueno")
- "negativo": queja, enojo, disconformidad, reclamo por incumplimiento (ej: "nunca pasan", "pésimo servicio", "no vinieron", "qué mal")
- "neutro": consulta, información, sin carga emocional clara (ej: "mi dirección es...", "cuándo pasan", saludos simples)

Mensaje: "${texto}"

Responde SOLO con JSON válido, sin texto adicional: {"sentimiento": "positivo" o "negativo" o "neutro"}`

        const response = await ollama.chat({
            model: 'llava:13b',
            messages: [{ role: 'user', content: pregunta }],
            options: { temperature: 0.1 }
        })

        const respuestaTexto = response.message.content
        const match = respuestaTexto.match(/\{[\s\S]*?\}/)
        if (match) {
            const json = JSON.parse(match[0])
            if (['positivo', 'negativo', 'neutro'].includes(json.sentimiento)) {
                return json.sentimiento
            }
        }
        return 'neutro'
    } catch (e) {
        console.error('Error analizando sentimiento:', e.message)
        return 'neutro'
    }
}

// Procesa el sentimiento de forma asíncrona, sin bloquear la respuesta del bot
function registrarSatisfaccion(numero, jid, texto, jidAlt) {
    // No registrar números de prueba en satisfaccion
    const numLimpio = numero.replace(/\D/g,'').replace(/^549/,'').replace(/^54/,'')
    const todosPrueba = [
        NUMERO_PRUEBA.replace(/^549/,''),
        ...NUMEROS_PRUEBA_EXTRA.map(n => n.replace(/^549/,''))
    ]
    if (todosPrueba.some(p => numLimpio === p.replace(/^549/,'') || numLimpio.endsWith(p.slice(-8)))) return
    // No bloquea: corre en background
    analizarSentimiento(texto).then(sentimiento => {
        try {
            const datos = cargarSatisfaccion()

            // Buscar dirección del pedido más reciente de este número, si existe
            const pedidos = cargarPedidos()
            const pedidoReciente = pedidos
                .filter(p => p.numero === numero)
                .sort((a, b) => new Date(b.fecha) - new Date(a.fecha))[0]
            // Solo usar dirección del pedido, no el mensaje libre (evita textos incorrectos)
            const direccion = pedidoReciente?.direccion || null
            const celularReal = jidANumero(numero, jidAlt)

            datos.push({
                numero,
                celular: celularReal, // vacío si no tenemos el numero real (no se inventa)
                direccion,
                sentimiento,
                mensaje: texto,
                fecha: new Date().toISOString()
            })
            guardarSatisfaccion(datos)

            const emoji = sentimiento === 'positivo' ? '🟢' : sentimiento === 'negativo' ? '🔴' : '⚪'
            console.log(`${emoji} [Satisfacción] +${numero}: ${sentimiento} - "${texto.slice(0, 50)}"`)
        } catch (e) {
            console.error('Error registrando satisfacción:', e.message)
        }
    })
}

// ===== LÓGICA PRINCIPAL =====
async function procesarMensaje(sock, msg) {
    const jid = msg.key.remoteJid
    const jidAlt = msg.key.remoteJidAlt || null // numero real si el remoteJid es @lid
    console.log(`📩 Mensaje de: ${jid}${jidAlt ? ' (real: ' + jidAlt + ')' : ''}`)
    if (jid.includes('@g.us')) return
    if (msg.key.fromMe) return

    // Cosecha de mapeo LID->número real (pendiente 3): si el chat es @lid y
    // tenemos el número real en remoteJidAlt, lo persistimos. Si no lo tenemos,
    // intentamos resolverlo activamente vía Baileys (en background).
    if (jid.includes('@lid')) {
        if (jidAlt && jidAlt.includes('@s.whatsapp.net')) registrarLidMap(jid, jidAlt)
        else if (!buscarPNenMap(jid)) resolverPNdeLID(sock, jid)
    }

    // Filtro modo prueba: 3814461809 siempre responde; otros números de prueba solo fuera de horario operativo
    if (MODO_PRUEBA) {
        // Resolver número real del JID (puede ser @lid sin número visible)
        const pnResuelto = buscarPNenMap(jid) || jid.replace(/@.*$/, '')
        const pnLimpio = pnResuelto.replace(/\D/g, '').replace(/^549/, '')
        const esNumeroPrincipal = jid.includes(NUMERO_PRUEBA) || jid.includes(JID_PRUEBA) ||
            pnLimpio === NUMERO_PRUEBA.replace(/^549/, '') ||
            pnLimpio === JID_PRUEBA
        const esNumeroPruebaHorario = esPruebaHorarioActivo(jid)
        if (!esNumeroPrincipal && !esNumeroPruebaHorario) return
    }

    const contenidoMsg = desempaquetarMensaje(msg.message)
    const texto = (contenidoMsg?.conversation
        || contenidoMsg?.extendedTextMessage?.text
        || '').trim()
    const tieneFoto = !!(contenidoMsg?.imageMessage)
    const numero = jid.replace('@s.whatsapp.net', '').replace('@lid', '')

    if (!estados[jid]) estados[jid] = { paso: 'inicio' }
    const estado = estados[jid]

    // ===== REINICIO DE CONVERSACIÓN =====
    // Si el vecino escribe palabras clave en cualquier momento → volver al menú
    const palabrasReinicio = /^(hola|menu|menú|inicio|cancelar|reiniciar|volver|empezar|comenzar|salir|0|00)$/i
    if (texto && palabrasReinicio.test(texto.trim()) && estado.paso !== 'inicio') {
        estados[jid] = { paso: 'inicio' }
        guardarEstados()
        await sock.sendMessage(jid, { text: MSG.bienvenida })
        return
    }

    // Análisis de satisfacción en background (solo vecinos, no personal interno)
    if (texto && texto.length >= 3 && !esPersonalInterno(numero)) {
        registrarSatisfaccion(numero, jid, texto, jidAlt)
        setTimeout(actualizarDashboardData, 3000) // actualizar dashboard 3 seg después
    }

    const ahora = new Date()
    const hora = ahora.getHours()
    const dia = ahora.getDay()
    const enHorario = dia >= 1 && dia <= 5 && hora >= 8 && hora < 13

    // ===== COMANDOS MANUALES DE ORDEN =====
    if (texto.toLowerCase() === 'generar orden') {
        await sock.sendMessage(jid, { text: '⚙️ Generando orden de servicio del día...' })
        await generarOrdenServicio(sock)
        return
    }

    if (texto.toLowerCase() === 'generar orden norte') {
        await sock.sendMessage(jid, { text: '⚙️ Generando orden de zona Norte...' })
        await generarOrdenServicioZona(sock, 'norte', new Date())
        return
    }

    if (texto.toLowerCase() === 'generar orden sur') {
        await sock.sendMessage(jid, { text: '⚙️ Generando orden de zona Sur...' })
        await generarOrdenServicioZona(sock, 'sur', new Date())
        return
    }

    // Comando: generar orden 2026-06-10
    const matchFecha = texto.match(/^generar orden (\d{4}-\d{2}-\d{2})$/i)
    if (matchFecha) {
        const fecha = new Date(matchFecha[1])
        const zona = zonaDelDia(fecha)
        if (!zona) {
            await sock.sendMessage(jid, { text: `⚠️ El ${matchFecha[1]} es feriado o domingo, no hay orden para ese día.` })
        } else {
            await sock.sendMessage(jid, { text: `⚙️ Generando orden de ${zona} para el ${matchFecha[1]}...` })
            await generarOrdenServicioZona(sock, zona, fecha)
        }
        return
    }

    // ===== COMANDO MANUAL: VER RESUMEN DE SATISFACCIÓN (solo prueba) =====
    if (texto.toLowerCase() === 'satisfaccion' || texto.toLowerCase() === 'satisfacción') {
        const datos = cargarSatisfaccion()
        const positivos = datos.filter(d => d.sentimiento === 'positivo').length
        const negativos = datos.filter(d => d.sentimiento === 'negativo').length
        const neutros = datos.filter(d => d.sentimiento === 'neutro').length
        await sock.sendMessage(jid, {
            text: `📊 *Resumen de satisfacción*\n\n🟢 Positivos: ${positivos}\n🔴 Negativos: ${negativos}\n⚪ Neutros: ${neutros}\n\nTotal registros: ${datos.length}\n\n_Detalle completo en satisfaccion.json_`
        })
        return
    }

    // ===== COMANDO MANUAL: GENERAR REPORTE DE MULTAS (solo prueba) =====
    if (texto.toLowerCase() === 'generar multas') {
        await sock.sendMessage(jid, { text: '⚙️ Generando reporte de multas...' })
        await generarReporteMultas(sock)
        await sock.sendMessage(jid, { text: '✅ Listo, revisá la carpeta del proyecto y el WhatsApp de destino.' })
        return
    }

    // ===== CONFIRMAR CUMPLIMIENTO =====
    if (estado.paso === 'esperandoConfirmacionCumplimiento') {
        const resp = texto.toLowerCase()
        if (['sí','si','s'].includes(resp) || detectaCumplimiento(texto)) {
            cerrarPedido(numero, 'completado')
            await sock.sendMessage(jid, { text: MSG.cumplido })
            estados[jid] = { paso: 'inicio' }
            guardarEstados()
        } else if (resp === 'no') {
            cerrarPedido(numero, 'incumplido')
            await sock.sendMessage(jid, { text: MSG.incumplido })
            if (GRUPO_ASESORES) {
                const p = getPedidoActivo(numero)
                await sock.sendMessage(GRUPO_ASESORES, {
                    text: `🚨 *Incumplimiento*\nNúmero: +${numero}\nTipo: ${p?.tipo}\nDirección: ${p?.mensaje}`
                })
            }
            estados[jid] = { paso: 'inicio' }
            guardarEstados()
        }
        return
    }

    // ===== VECINO CON PEDIDO ACTIVO =====
    const pedidoActivo = getPedidoActivo(numero)
    if (pedidoActivo && estado.paso === 'inicio') {
        if (detectaCumplimiento(texto)) {
            cerrarPedido(numero, 'completado')
            await sock.sendMessage(jid, { text: MSG.cumplido })
            guardarEstados()
            return
        }
        await sock.sendMessage(jid, { text: MSG.pedidoActivo })
        estados[jid] = { paso: 'esperandoOpcionPedidoActivo' }
        guardarEstados()
        return
    }

    if (estado.paso === 'esperandoOpcionPedidoActivo') {
        if (texto === '1') {
            await sock.sendMessage(jid, { text: MSG.bienvenida })
            estados[jid] = { paso: 'esperandoOpcion' }
        } else if (texto === '2') {
            const p = pedidoActivo
            await sock.sendMessage(jid, {
                text: MSG.verPedido(p.tipo, p.mensaje, new Date(p.fechaOrden), new Date(p.fechaVencimiento))
            })
            estados[jid] = { paso: 'inicio' }
        } else if (texto === '3') {
            if (!enHorario) {
                await sock.sendMessage(jid, { text: MSG.fueraDeHorario })
            } else {
                await sock.sendMessage(jid, { text: MSG.opcion6 })
                if (GRUPO_ASESORES) {
                    const celularReal3 = jidANumero(jid, null)
                    const displayNum3 = celularReal3 || `+${numero}`
                    await sock.sendMessage(GRUPO_ASESORES, {
                        text: `🔔 *Asesor solicitado*\n📱 ${displayNum3}\nPedido: ${pedidoActivo.tipo}`
                    })
                }
            }
            estados[jid] = { paso: 'inicio' }
        }
        guardarEstados()
        return
    }

    // ===== NO CORRESPONDE =====
    if (texto && noCorresponde(texto) && estado.paso !== 'esperandoDatos') {
        await sock.sendMessage(jid, { text: MSG.noCorresponde })
        estados[jid] = { paso: 'inicio' }
        guardarEstados()
        return
    }

    // ===== MENÚ PRINCIPAL =====
    if (estado.paso === 'inicio') {
        // Si hay un estado guardado anterior con tipo (conversación incompleta)
        if (estado.tipoAnterior) {
            await sock.sendMessage(jid, { text: MSG.conversacionAnterior(estado.tipoAnterior) })
            estados[jid] = { paso: 'esperandoOpcionConversacion', tipoAnterior: estado.tipoAnterior, estadoAnterior: estado.estadoAnterior }
            guardarEstados()
            return
        }
        await sock.sendMessage(jid, { text: MSG.bienvenida })
        estados[jid] = { paso: 'esperandoOpcion' }
        guardarEstados()
        return
    }

    // ===== OPCIÓN CONVERSACIÓN ANTERIOR =====
    if (estado.paso === 'esperandoOpcionConversacion') {
        if (texto === '1') {
            // Nueva consulta
            await sock.sendMessage(jid, { text: MSG.bienvenida })
            estados[jid] = { paso: 'esperandoOpcion' }
        } else if (texto === '2') {
            // Continuar con lo anterior
            if (estado.estadoAnterior) {
                estados[jid] = estado.estadoAnterior
                await sock.sendMessage(jid, { text: `Continuamos 👍 ¿En qué te podemos ayudar con *${estado.tipoAnterior}*?` })
            } else {
                await sock.sendMessage(jid, { text: MSG.bienvenida })
                estados[jid] = { paso: 'esperandoOpcion' }
            }
        } else {
            await sock.sendMessage(jid, { text: MSG.conversacionAnterior(estado.tipoAnterior) })
        }
        guardarEstados()
        return
    }

    if (estado.paso === 'esperandoOpcion') {
        const op = texto
        if (op === '1') {
            await sock.sendMessage(jid, { text: MSG.opcion1 })
            estados[jid] = { paso: 'esperandoSubtipoRNH', tipoAnterior: 'Residuos no habituales' }
        } else if (op === '2') {
            await sock.sendMessage(jid, { text: MSG.opcion2 })
            estados[jid] = { paso: 'esperandoDatos', tipo: 'Basural', tipoAnterior: 'Basural o volcadero' }
        } else if (op === '3') {
            await sock.sendMessage(jid, { text: MSG.opcion3 })
            estados[jid] = { paso: 'esperandoDatos', tipo: 'Barrido', tipoAnterior: 'Barrido' }
        } else if (op === '4') {
            await sock.sendMessage(jid, { text: MSG.opcion4 })
            estados[jid] = { paso: 'esperandoDatos', tipo: 'Falta de recolección', tipoAnterior: 'Falta de recolección' }
        } else if (op === '5') {
            await sock.sendMessage(jid, { text: MSG.opcion5 })
            estados[jid] = { paso: 'esperandoTipoPersona', tipoAnterior: 'Persona o vehículo arrojando basura' }
        } else if (op === '6') {
            if (!enHorario) {
                await sock.sendMessage(jid, { text: MSG.fueraDeHorario })
            } else {
                await sock.sendMessage(jid, { text: MSG.opcion6 })
                if (GRUPO_ASESORES) {
                    const celularReal6 = jidANumero(jid, null)
                    const displayNum6 = celularReal6 || `+${numero}`
                    await sock.sendMessage(GRUPO_ASESORES, {
                        text: `🔔 *Asesor solicitado*\n📱 ${displayNum6}\nHora: ${ahora.toLocaleTimeString('es-AR')}`
                    })
                }
            }
            estados[jid] = { paso: 'inicio' }
        } else {
            await sock.sendMessage(jid, { text: MSG.noEntiendo })
        }
        guardarEstados()
        return
    }

    if (estado.paso === 'esperandoSubtipoRNH') {
        const op = texto
        const subtipos = {
            '1': { msg: MSG.opcion1_madera, tipo: 'RNH - Madera/Poda' },
            '2': { msg: MSG.opcion1_escombros, tipo: 'RNH - Escombros' },
            '3': { msg: MSG.opcion1_muebles, tipo: 'RNH - Muebles/Electrodomésticos' },
            '4': { msg: MSG.opcion1_otro, tipo: 'RNH - Otro' }
        }
        if (subtipos[op]) {
            await sock.sendMessage(jid, { text: subtipos[op].msg })
            estados[jid] = { paso: 'esperandoDatos', tipo: subtipos[op].tipo }
        } else {
            await sock.sendMessage(jid, { text: MSG.opcion1 })
        }
        guardarEstados()
        return
    }

    if (estado.paso === 'esperandoTipoPersona') {
        if (texto === '1') {
            await sock.sendMessage(jid, { text: MSG.opcion5a })
            estados[jid] = { paso: 'esperandoDatos', tipo: 'Vehículo arrojando basura' }
        } else if (texto === '2') {
            await sock.sendMessage(jid, { text: MSG.opcion5b })
            estados[jid] = { paso: 'esperandoDatos', tipo: 'Persona arrojando basura' }
        } else {
            await sock.sendMessage(jid, { text: MSG.opcion5 })
        }
        guardarEstados()
        return
    }

    // ===== RECEPCIÓN DE DATOS =====
    if (estado.paso === 'esperandoDatos') {
        if (tieneFoto) {
            try {
                await sock.sendMessage(jid, { text: '🔍 Analizando tu foto...' })
                const buffer = await downloadMediaMessage(msg, 'buffer', {})
                const analisis = await analizarFoto(buffer, estado.tipo)

                if (!analisis.esValida) {
                    await sock.sendMessage(jid, { text: MSG.fotoInvalida(analisis.razon) })
                    estados[jid] = { ...estado, fotoRecibida: false }
                    guardarEstados()
                    return
                }

                // Si ya teníamos una dirección pendiente de un mensaje anterior
                const direccionFinal = texto && pareceDireccion(texto) ? texto : estado.direccionPendiente

                // Foto válida
                if (direccionFinal) {
                    await confirmarPedido(sock, jid, numero, estado.tipo, { mensaje: direccionFinal, tieneFoto: true })
                } else {
                    // Tiene foto válida pero falta dirección (o el texto no parece dirección)
                    estados[jid] = { ...estado, fotoRecibida: true }
                    if (estado.tipo === 'Vehículo arrojando basura') {
                        await sock.sendMessage(jid, { text: MSG.fotoRecibidaSinDireccionVehiculo })
                    } else if (estado.tipo === 'Persona arrojando basura') {
                        await sock.sendMessage(jid, { text: MSG.fotoRecibidaSinDireccionPersona })
                    } else {
                        await sock.sendMessage(jid, { text: MSG.fotoRecibidaSinDireccion })
                    }
                    guardarEstados()
                }

            } catch (e) {
                console.error('Error foto:', e)
                await sock.sendMessage(jid, { text: MSG.fotoRecibidaSinDireccion })
                estados[jid] = { ...estado, fotoRecibida: false }
                guardarEstados()
            }

        } else if (texto && estado.fotoRecibida) {
            // Ya tenía foto válida, ahora espera dirección
            if (!pareceDireccion(texto)) {
                await sock.sendMessage(jid, {
                    text: 'No pudimos identificar una dirección válida 📍\n\nPor favor enviá calle y número (o calle y calle) para poder gestionar tu pedido.'
                })
                guardarEstados()
                return
            }
            await confirmarPedido(sock, jid, numero, estado.tipo, { mensaje: texto, tieneFoto: true })

        } else if (texto) {
            // Solo texto sin foto todavía
            if (!pareceDireccion(texto)) {
                // Verificar si pasó tiempo desde el último mensaje (más de 30 min = conversación nueva)
                const ultimoMsg = estado.ultimoMensaje ? new Date(estado.ultimoMensaje) : null
                const minutosDesde = ultimoMsg ? (Date.now() - ultimoMsg.getTime()) / 60000 : 999
                const tipoAnterior = estado.tipoAnterior || estado.tipo || 'tu pedido anterior'

                if (minutosDesde > 30 || !ultimoMsg) {
                    // Pasó tiempo — preguntar si quiere continuar o empezar de nuevo
                    await sock.sendMessage(jid, {
                        text: `Hola 👋 Tenemos una conversación anterior sobre *${tipoAnterior}* que quedó pendiente.\n\n¿Qué querés hacer?\n\n1️⃣ Continuar con el pedido anterior\n2️⃣ Iniciar una consulta nueva`
                    })
                    estados[jid] = { ...estado, paso: 'esperandoOpcionContinuar', ultimoMensaje: new Date().toISOString() }
                } else {
                    // Conversación reciente — pedir los datos que faltan
                    await sock.sendMessage(jid, {
                        text: `Para completar tu pedido necesitamos la *foto* 📸 y la *dirección* 📍\n\n¿Podés enviarlas?`
                    })
                    estados[jid] = { ...estado, ultimoMensaje: new Date().toISOString() }
                }
                guardarEstados()
                return
            }
            // Tiene formato de dirección pero falta la foto
            estados[jid] = { ...estado, direccionPendiente: texto, ultimoMensaje: new Date().toISOString() }
            await sock.sendMessage(jid, {
                text: `Gracias, registramos la dirección: *${texto}* 📍\n\nAhora necesitamos que envíes la *foto* para completar tu pedido 📸`
            })
            guardarEstados()
        }
        return
    }

    // ===== OPCIÓN CONTINUAR / NUEVA CONVERSACIÓN =====
    if (estado.paso === 'esperandoOpcionContinuar') {
        if (texto === '1') {
            // Continuar con el pedido anterior
            const tipoAnterior = estado.tipoAnterior || estado.tipo || 'pedido'
            await sock.sendMessage(jid, {
                text: `Continuamos con tu pedido de *${tipoAnterior}* 👍\n\nPor favor enviá la *foto* 📸 y la *dirección* 📍`
            })
            estados[jid] = { ...estado, paso: 'esperandoDatos', ultimoMensaje: new Date().toISOString() }
        } else if (texto === '2') {
            // Iniciar nueva conversación
            await sock.sendMessage(jid, { text: MSG.bienvenida })
            estados[jid] = { paso: 'esperandoOpcion', ultimoMensaje: new Date().toISOString() }
        } else {
            await sock.sendMessage(jid, {
                text: `Por favor respondé:\n1️⃣ Continuar con el pedido anterior\n2️⃣ Iniciar una consulta nueva`
            })
        }
        guardarEstados()
        return
    }
}

// ===== INICIO =====
async function iniciarBot() {
    cargarEstados()

    // ===== PARSER CSV correcto (maneja comillas y comas internas) =====
    function parsearLineaCSV(linea) {
        const resultado = []
        let actual = ''
        let dentroComillas = false
        for (let i = 0; i < linea.length; i++) {
            const c = linea[i]
            if (c === '"') {
                if (dentroComillas && linea[i+1] === '"') {
                    actual += '"'
                    i++
                } else {
                    dentroComillas = !dentroComillas
                }
            } else if (c === ',' && !dentroComillas) {
                resultado.push(actual)
                actual = ''
            } else {
                actual += c
            }
        }
        resultado.push(actual)
        return resultado
    }

    function parsearCSVCompleto(contenido) {
        // Maneja saltos de línea dentro de campos entrecomillados
        const filas = []
        let actual = ''
        let dentroComillas = false
        for (let i = 0; i < contenido.length; i++) {
            const c = contenido[i]
            if (c === '"') dentroComillas = !dentroComillas
            if ((c === '\n' || c === '\r') && !dentroComillas) {
                if (actual.trim()) filas.push(actual)
                actual = ''
                if (c === '\r' && contenido[i+1] === '\n') i++
            } else {
                actual += c
            }
        }
        if (actual.trim()) filas.push(actual)
        return filas
    }

    // Si satisfaccion.json está vacío O solo tiene mensajes inventados, importar desde base_electoral.csv
    const satActual = cargarSatisfaccion()
    const satValido = satActual.filter(s => {
        const msg = s.mensaje || ''
        return !msg.match(/^Hola!\s+(Te escribimos|Somos del equipo)/i)
    })
    // Solo reimportar si está completamente vacío (no sobreescribir datos válidos)
    if (satActual.length === 0 && fs.existsSync('base_electoral.csv')) {
        try {
            // Columnas reales del CSV: numero_envio,id_interno,nombre,direccion,sentimiento,canal,ultimo_contacto,mensaje_contacto
            const csv = fs.readFileSync('base_electoral.csv', 'utf8')
            const filas = parsearCSVCompleto(csv).slice(1) // saltar encabezado
            const lidMapActual = cargarLidMap()
            const datos = filas.map(l => {
                const p = parsearLineaCSV(l)
                const numeroEnvio = (p[0] || '').trim()
                const idInterno = (p[1] || '').trim()
                const nombre = (p[2] || '').trim()
                const direccion = (p[3] || '').trim()
                const sentimiento = (p[4] || 'neutro').trim()
                const canal = (p[5] || 'WhatsApp Ambiente').trim()
                const fecha = (p[6] || new Date().toISOString()).trim()
                const mensajeRaw = (p[7] || '').trim()
                // NO guardar mensajes inventados por el clasificador
                const mensaje = /^Hola!\s+(Te escribimos|Somos del equipo)/i.test(mensajeRaw) ? '' : mensajeRaw

                // Intentar resolver celular desde lid_map usando idInterno
                let celular = numeroEnvio
                if (!celular && idInterno) {
                    const pnResuelto = lidMapActual[idInterno]
                    if (pnResuelto) celular = jidANumero(idInterno + '@lid', pnResuelto + '@s.whatsapp.net')
                }

                return {
                    numero: idInterno || numeroEnvio,
                    celular: celular || '',
                    nombre, direccion, sentimiento, canal, fecha, mensaje
                }
            }).filter(d => d.numero)
            guardarSatisfaccion(datos)
            console.log(`📊 Cargados ${datos.length} registros históricos de base_electoral.csv`)
            actualizarDashboardData()
        } catch(e) {
            console.error('Error cargando base_electoral.csv:', e.message)
        }
    }
    const { state, saveCreds } = await useMultiFileAuthState('auth_info')
    const { version } = await fetchLatestBaileysVersion()

    const sock = makeWASocket({
        version, auth: state, printQRInTerminal: false
    })

    sock.ev.on('creds.update', saveCreds)

    // Pendiente 3: WhatsApp avisa mapeos LID->número mediante este evento.
    // Los persistimos para enriquecer dashboard y base electoral.
    sock.ev.on('lid-mapping.update', (mapping) => {
        try {
            const pares = Array.isArray(mapping) ? mapping : [mapping]
            let nuevos = 0
            for (const { lid, pn } of pares) {
                if (registrarLidMap(lid, pn)) nuevos++
            }
            if (nuevos > 0) console.log(`🔗 ${nuevos} mapeo(s) LID→número aprendidos`)
        } catch (e) {
            console.error('Error en lid-mapping.update:', e.message)
        }
    })

    sock.ev.on('connection.update', async (update) => {
        const { connection, qr } = update
        if (qr) qrcode.generate(qr, { small: true })
        if (connection === 'open') {
            console.log('✅ Bot conectado — MODO PRUEBA ACTIVO')
            console.log(`📱 Solo responde a: ${NUMERO_PRUEBA} / ${JID_PRUEBA} / ${NUMEROS_PRUEBA_EXTRA.join(' / ')}`)
            // Pre-cargar JIDs de prueba desde lid_map si existen
            const lidMapActual = cargarLidMap()
            for (const [lid, pn] of Object.entries(lidMapActual)) {
                const pnLimpio = pn.replace(/\D/g, '').replace(/^549/, '')
                if (NUMEROS_PRUEBA_HORARIO.some(n => n.replace(/\D/g,'').replace(/^549/,'') === pnLimpio)) {
                    console.log(`✅ JID de prueba cargado: ${lid} → +${pn}`)
                }
            }

            // Identificar grupo BAM
            try {
                const grupos = await sock.groupFetchAllParticipating()
                for (const [id, info] of Object.entries(grupos)) {
                    if (info.subject && info.subject.toUpperCase() === NOMBRE_GRUPO_BAM) {
                        GRUPO_BAM_ID = id
                        console.log(`✅ Grupo BAM identificado: ${id} (escuchando en vivo)`)
                    }
                }
                if (!GRUPO_BAM_ID) console.log('⚠️  Grupo BAM no encontrado, se reintentará detección por mensajes')
            } catch (e) {
                console.error('Error buscando grupo BAM:', e.message)
            }

            // Verificación de vencimientos cada 30 min
            setInterval(() => verificarVencimientos(sock), 30 * 60 * 1000)

            // Consultar API de reclamos al arrancar
            actualizarReclamos()

            // Scheduler de orden de servicio diaria a las 12:50hs
            programarGeneracionOrden(sock)
            console.log('🕐 Generación automática de orden de servicio programada: 12:50hs')
        }
        if (connection === 'close') {
            console.log('🔄 Reconectando...')
            iniciarBot()
        }
    })

    // Analizar historial para detectar cumplimientos ya ocurridos
    sock.ev.on('messaging-history.set', async ({ messages }) => {
        const pedidos = cargarPedidos()
        let actualizados = 0

        // Cosecha de mapeos LID->número desde el historial (pendiente 3)
        let lidNuevos = 0
        for (const msg of messages) {
            const rj = msg.key.remoteJid
            const alt = msg.key.remoteJidAlt
            if (rj?.includes('@lid') && alt?.includes('@s.whatsapp.net')) {
                if (registrarLidMap(rj, alt)) lidNuevos++
            }
        }
        if (lidNuevos > 0) console.log(`🔗 ${lidNuevos} mapeo(s) LID→número aprendidos del historial`)

        for (const msg of messages) {
            if (msg.key.fromMe) continue
            if (msg.key.remoteJid?.includes('@g.us')) continue

            const texto = (msg.message?.conversation
                || msg.message?.extendedTextMessage?.text
                || '').trim()

            if (!texto || !detectaCumplimiento(texto)) continue

            const numero = msg.key.remoteJid
                .replace('@s.whatsapp.net', '')
                .replace('@lid', '')

            if (esPersonalInterno(numero)) continue

            const pedido = pedidos.find(p =>
                p.numero === numero &&
                p.estado === 'pendiente' &&
                msg.messageTimestamp &&
                new Date(Number(msg.messageTimestamp) * 1000) >= new Date(p.fecha)
            )

            if (pedido) {
                pedido.estado = 'completado'
                pedido.fechaCierre = new Date(Number(msg.messageTimestamp) * 1000).toISOString()
                pedido.cierreDetectadoEn = texto
                actualizados++
            }
        }

        if (actualizados > 0) {
            guardarPedidos(pedidos)
            console.log(`✅ ${actualizados} pedidos marcados como completados desde el historial`)
        }

        // Backfill del grupo BAM: capturar mensajes del historial que no se
        // hayan procesado en vivo (bot offline, mensajes temporales, etc.).
        // procesarMensajeBAM es idempotente gracias al dedupe por msg.key.id.
        if (GRUPO_BAM_ID) {
            let bamBackfill = 0
            for (const msg of messages) {
                if (msg.key.fromMe) continue
                if (msg.key.remoteJid !== GRUPO_BAM_ID) continue
                try {
                    const antes = cargarPedidos().length + cargarMultas().length
                    await procesarMensajeBAM(sock, msg)
                    const despues = cargarPedidos().length + cargarMultas().length
                    if (despues > antes) bamBackfill++
                } catch (e) {
                    console.error('Error backfill BAM:', e.message)
                }
            }
            if (bamBackfill > 0) console.log(`📋 ${bamBackfill} mensajes BAM capturados desde el historial`)
        }
    })

    sock.ev.on('messages.upsert', async ({ messages, type }) => {
        for (const msg of messages) {
            try {
                // Mensajes del grupo BAM -> procesar como pedido.
                // Se capturan para CUALQUIER type (notify/append/etc.): así no se
                // pierden los que llegan como historial o estando el bot offline.
                // El dedupe por msg.key.id evita procesar dos veces.
                if (msg.key.remoteJid === GRUPO_BAM_ID) {
                    await procesarMensajeBAM(sock, msg)
                    continue
                }
                // Detectar grupo BAM dinámicamente si aún no se identificó
                if (!GRUPO_BAM_ID && msg.key.remoteJid?.includes('@g.us')) {
                    try {
                        const meta = await sock.groupMetadata(msg.key.remoteJid)
                        if (meta.subject && meta.subject.toUpperCase() === NOMBRE_GRUPO_BAM) {
                            GRUPO_BAM_ID = msg.key.remoteJid
                            console.log(`✅ Grupo BAM identificado dinámicamente: ${GRUPO_BAM_ID}`)
                            await procesarMensajeBAM(sock, msg)
                            continue
                        }
                    } catch (e) {}
                }

                // Mensajes de grupos que no son BAM -> ignorar
                if (msg.key.remoteJid?.includes('@g.us')) continue

                // Vecinos: SOLO mensajes en vivo (notify). Nunca responder a
                // historial/append, para no contestar mensajes viejos.
                if (type !== 'notify') continue
                await procesarMensaje(sock, msg)
            } catch (e) {
                console.error('Error:', e)
            }
        }
    })
}

// Auto-arranque solo al ejecutar directamente (node index.js).
// Si el módulo se requiere (p.ej. para tests), no conecta a WhatsApp.
if (require.main === module) {
    iniciarBot()
}

module.exports = {
    leerItemsSheets, parsearSheetMarkdown, extraerDireccionBAM, desempaquetarMensaje,
    jidANumero, registrarLidMap, buscarPNenMap, formatearNumeroLocal, soloDigitos
}
