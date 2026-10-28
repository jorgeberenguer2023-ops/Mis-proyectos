# CodeForge Mobile

MVP inicial de editor de código para Android/iOS con Expo, explorador de archivos básico, catálogo de 35 lenguajes, edición en WebView y almacenamiento local.

## Requisitos
- Node.js LTS
- Android Studio para generar/ejecutar Android localmente
- Xcode en macOS para compilar iOS

## Arranque
```bash
npm install
npx expo start
```
Para Android nativo:
```bash
npx expo run:android
```
Para iOS (macOS):
```bash
npx expo run:ios
```

## Qué incluye este MVP
- Editor con números de línea y resaltado básico de sintaxis dentro de una WebView.
- Selección de lenguaje para 35 lenguajes y formatos.
- Crear, cambiar de archivo y guardar localmente proyectos de demostración.
- Persistencia de documentos mediante AsyncStorage.
- Pantalla GitHub de demostración y guía de seguridad para la futura integración.

## Límites actuales
La lista de lenguajes define reconocimiento y modo de edición; no implica que todos tengan resaltado avanzado, LSP, compilador o ejecución. Este MVP aún no implementa OAuth real, clonación Git, push/pull ni ejecución de código. La autenticación GitHub debe implementarse con OAuth App/GitHub App y un backend seguro cuando sea necesario; no se debe incluir un client secret en la app móvil.

## Próximos hitos
1. Sustituir el resaltado básico por CodeMirror 6 y sus paquetes de lenguaje empaquetados localmente.
2. Añadir gestión real de archivos del dispositivo y un modelo de proyectos.
3. Implementar GitHub App/OAuth con PKCE o backend, scopes mínimos y tokens almacenados de forma segura.
4. Implementar Git real con biblioteca compatible o backend y resolución de conflictos.
5. Añadir runners aislados por lenguaje; nunca ejecutar código arbitrario en el servidor principal.

## Ejecución de código en el MVP

La versión 0.3 añade ejecución local básica para JavaScript en el WebView del editor: pulsa **▶ Ejecutar** y se muestra la salida de `console.log`, advertencias y errores. No es Node.js, no soporta APIs de Node ni módulos, y no es un sandbox de seguridad; ejecuta únicamente código propio y de confianza.

Python, TypeScript y los demás lenguajes todavía no se ejecutan. Para ellos hace falta integrar intérpretes específicos o un servicio remoto de ejecución aislado, autenticado y con límites de recursos. No introduzcas claves, tokens ni código de terceros en el ejecutor.
