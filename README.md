# 📇 Registro de Negocios

Aplicación de escritorio para gestionar un directorio de clientes/negocios por rubro — pensada para pequeños negocios que necesitan organizar contactos, hacer seguimiento comercial y mantener todo en un solo lugar, **sin depender de internet ni de servicios en la nube**.

Hecha en **Python + Tkinter**, con base de datos **SQLite** local. Corre como script `.py` o como ejecutable `.exe` portable en Windows.

---

## 📌 Índice

- [¿Qué problema resuelve?](#-qué-problema-resuelve)
- [Funciones](#-funciones)
- [Cómo funciona por dentro](#-cómo-funciona-por-dentro)
- [Atajos de teclado](#-atajos-de-teclado)
- [Instalación y uso](#-instalación-y-uso)
- [Generar el .exe](#-generar-el-exe)
- [Estructura de archivos que genera el programa](#-estructura-de-archivos-que-genera-el-programa)
- [Seguridad y respaldo de datos](#-seguridad-y-respaldo-de-datos)
- [Roadmap / a futuro](#-roadmap--a-futuro)
- [Publicar en LinkedIn / Instagram](#-publicar-en-linkedin--instagram)

---

## 🎯 ¿Qué problema resuelve?

Cuando un negocio empieza a tener decenas o cientos de contactos (clientes, proveedores, prospectos), las hojas de cálculo sueltas se vuelven difíciles de mantener: no hay historial, no hay alertas de seguimiento, es fácil duplicar registros y no hay forma rápida de buscar "¿ya tengo el teléfono de este negocio guardado?".

Esta app centraliza todo eso en un programa liviano, de un solo archivo, que cualquiera en el equipo puede abrir sin instalar nada raro — y que guarda los datos cifrados y con respaldos automáticos.

---

## 🧩 Funciones

### Gestión de registros

| Función | Descripción |
|---|---|
| Registro por rubro | Cada negocio pertenece a un rubro (Cafetería, Ferretería, etc.) y recibe un código automático (`PREFIJO + número`, ej. `TSCA01`). |
| Editar / eliminar | Edición completa con validaciones, eliminación con confirmación. |
| Mover a otro rubro | Clic derecho sobre un registro → "Mover a otro rubro" (individual o en selección múltiple). El código conserva su prefijo original. |
| Duplicar registro | Clic derecho → "Duplicar", para dar de alta un negocio parecido sin escribir todo de nuevo. |
| Edición masiva | Selecciona varios registros y cambia un campo (ej. estado) en todos a la vez. |
| Favoritos ⭐ | Marca los negocios más importantes para encontrarlos rápido. |
| Etiquetas | Etiquetas libres por negocio, además del rubro. |
| Notas | Campo de notas/comentarios por registro. |
| Fecha de registro e historial | Cada negocio guarda su fecha de alta; todos los cambios (creación, edición, cambio de rubro, eliminación, importaciones) quedan en un historial consultable. |
| Seguimientos | Fecha de "próximo contacto" con recordatorio visual en el Dashboard cuando está vencido o es hoy. |
| Adjuntar archivos 📎 | Adjunta contratos, fotos o cotizaciones a cualquier negocio (clic derecho → Adjuntos). Se guardan como copia local en `adjuntos/`. |

### Organización y búsqueda

| Función | Descripción |
|---|---|
| Pestañas por rubro | Además de una "Vista general" con todos los negocios. |
| Búsqueda | Por nombre, código, rubro, contacto, ciudad, etiquetas, **teléfono** (ignora espacios/guiones/`+51`), correo, dirección, web y RUC. |
| Filtros | Por estado (Activo/Potencial/Inactivo) y otros criterios visibles en la barra de filtros. |
| Acciones rápidas | Clic derecho: copiar teléfono/correo, abrir WhatsApp, redactar en Gmail, abrir sitio web. |

### Importación

| Fuente | Cómo funciona |
|---|---|
| **Google Maps Scraper (CSV)** | Reconoce columnas `Name, Phone, Email, Website, Address, Category…` (ignora Instagram/Facebook/PlaceID/coordenadas/etc. porque no aportan al directorio). Muestra una **vista previa** antes de importar: puedes editar el nombre del rubro y las siglas del código, ver cuántas filas se omitirán y confirmar. Filas corruptas se saltan solas, sin frenar el resto. |
| **Base de Datos Previa** | Excel con una hoja por rubro, o CSV con columna de Rubro — el mismo formato que exporta este programa. |
| Deduplicación | Si un negocio con el mismo nombre (sin distinguir mayúsculas/espacios) ya existe, no se crea un duplicado — se omite y queda anotado en 🔔 Notificaciones para revisar. |

### Exportación e impresión

| Función | Descripción |
|---|---|
| Excel (.xlsx) | Con una hoja "Resumen" y una hoja por rubro. |
| CSV | Formato plano, compatible con cualquier hoja de cálculo. |
| PDF | Listado con tablas por rubro, formateado para imprimir. |
| Solo resultados filtrados | Cualquiera de los tres formatos puede exportar solo lo que está filtrado/buscado en pantalla. |
| Imprimir 🖨️ | Genera el PDF y lo manda directo a la impresora o visor predeterminado. |

### Dashboard (página aparte, con navegación propia)

| Función | Descripción |
|---|---|
| Estadísticas generales | Total de negocios, favoritos, seguimientos vencidos/hoy, negocios por rubro y por estado. |
| Actividad reciente | Últimos cambios registrados en el historial. |
| Gráficos avanzados 📈 | Pastel de negocios por rubro y barras de negocios agregados por mes (requiere `matplotlib`; si no está instalado, se muestra un aviso en vez de romperse). |

### Seguridad y respaldo

| Función | Descripción |
|---|---|
| 🔒 Cifrado en reposo | Al cerrar el programa correctamente, la base de datos se cifra (`negocios.db.enc`); se descifra automáticamente al volver a abrir. No requiere contraseña — la clave vive en `negocios.key` junto al programa. |
| 💾 Respaldos automáticos | Uno diario y otro al cerrar el programa; se conservan los últimos 20 en `backups/`. |
| ♻️ Restaurar respaldo | Desde ⚙ Ajustes, con respaldo de seguridad automático antes de restaurar. |

### Interfaz

| Función | Descripción |
|---|---|
| Diseño moderno | Paleta clara con acentos de color, tarjetas, botones con hover. |
| 🌙 Tema oscuro / ☀️ claro | Reconstruye toda la interfaz al cambiar. |
| 🎨 Color de acento personalizable | Varios colores predefinidos o un color a elección (selector de color del sistema). |
| Redimensionable | Panel central con divisor arrastrable + formulario con scroll (nunca se ocultan campos). |
| Recuerda tamaño/posición de ventana | Se guarda en `config.json`. |
| Atajos de teclado | Ver tabla más abajo. |

---

## ⚙️ Cómo funciona por dentro

- **Interfaz:** `tkinter` + `ttk`, con componentes propios (`ModernButton`, `ScrollableFrame`, `Card`) para lograr un look moderno sin librerías externas de UI.
- **Datos:** `SQLite` en modo `WAL` (mejor rendimiento con lecturas/escrituras concurrentes), con índices en las columnas más consultadas. Las tablas principales son `categories` (rubros), `businesses` (negocios), `history` (historial de cambios) y `attachments` (adjuntos).
- **Rendimiento:** al registrar, editar o mover un negocio, la interfaz **no reconstruye toda la tabla** — solo actualiza el contenido de las pestañas ya existentes, salvo que cambie la lista de rubros.
- **Exportación:** `openpyxl` (Excel) y `reportlab` (PDF).
- **Cifrado:** `cryptography` (Fernet, AES simétrico) sobre el archivo completo de la base de datos.
- **Gráficos:** `matplotlib` embebido en la ventana con `FigureCanvasTkAgg`.
- **Todo el código vive en un solo archivo (`app.py`)** para que sea fácil de compilar a `.exe` con PyInstaller sin dependencias sueltas que se puedan perder.

---

## ⌨️ Atajos de teclado

| Atajo | Acción |
|---|---|
| `Ctrl + Enter` | Registrar negocio / guardar edición |
| `Enter` (dentro de un campo) | Saltar al siguiente campo (o enviar si es el último) |
| `Ctrl + N` | Foco rápido en el selector de rubro |
| `Ctrl + E` | Editar el negocio seleccionado |
| `Ctrl + D` | Duplicar el negocio seleccionado |
| `Ctrl + F` | Ir al buscador |
| `Supr` | Eliminar el/los negocio(s) seleccionado(s) |
| `Ctrl + S` | Exportar |
| `F5` | Actualizar |
| `Esc` | Limpiar mensaje / cerrar diálogos |
