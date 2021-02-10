# Exportación de listados a documentos (Excel, pdf...)
Ejemplo realizado para productos Legacy (WebForms / ASP MVC 5 "clásico"), en donde se puede exportar una lista a un documento Excel

Los puntos tratados en el ejemplo para MVC, en el proyecto Export.Mvc son:
* MVC Excel: exportación a Excel, mediante un ActionResult personalizado (ExcelResult). Se puede proporcionar una lista fuertemente tipada o una dinámica. El ejemplo tiene ambas.
* MVC: Creación de un documento en formato PDF

Los puntos tratados en el ejemplo para Webforms, en el proyecto Export.WebForms son:
* MVC Excel: exportación a Excel, mediante dos botones en el formulario Excel.aspx. Se puede proporcionar una lista fuertemente tipada o una dinámica. El ejemplo tiene ambas.
* MVC: Creación de un documento en formato PDF

Los puntos tratados en el ejemplo tanto para MVC como Webforms son, mediante el proyecto Export.Common son:
* Excel Uso de la librería ClosedXML

## Comenzando 🚀

Puedes poner como punto de partida, cualquiera de estos proyectos:
* Export.Mvc
* Export.WebForms

### Pre-requisitos 📋

Como herramientas de desarrollo necesitarás:
* Visual Studio 2019, con el framework 4.5.2 (sí, es legacy)

```
Nota: Visual Studio Code también valdría
```

### Instalación 🔧

* No requiere nada especial

## Ejecutando la aplicación ⚙️

Si todo ha ido bien, tanto en MVC como en Webforms deberían aparcerer las páginas para poder navegar, no hay más misterio.

### ¿Qué falta? 🔩

Muchísimas cosas. Por ejemplo: 
* La clase Export.Common.ExcelGenerator, no admite clases complejas (con clases por debajo).
* Hacer un ejemplo completo en PDF con iTextSharp, y migrarlo de MVC a la librería común, para poder utilizarlo en WebForms
Si te animas a evolucionarlo... ningún problema ¡todo lo contrario! :)

## Construido con 🛠️

* [Microsoft Visual Studio Community 2019](https://visualstudio.microsoft.com/es/vs/) - IDE  de desarrollo
* [Resharper](https://www.jetbrains.com/es-es/resharper/) - Extensión del IDE VS para optimizar el desarrollo

## Versionado 📌

Usado [Git](https://git-scm.com//) para el versionado. Por ahora, no existen diferentes versiones disponibles. Si en un futuro existieran, estarían en los diferentes tags que se crearían.

## Autores ✒️

* **David Santesteban** - *Trabajo Inicial* - [davidsantes](https://github.com/davidsantes)

## Agradecimientos 🎁

* Plataforma de aprendizaje CampusMVP [CampusMvp](https://www.campusmvp.es/)
* A cualquiera que me invite a una cerveza 🍺. 
