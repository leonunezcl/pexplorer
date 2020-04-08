# Programa
Project Explorer - Analizador de código fuente proyectos Visual Basic 4,5 y 6.

# Autor
Luis Leonardo Nuñez Ibarra. Año 2000 - 2003. email : leo.nunez@gmail.com. 

Chileno, casado , tengo 2 hijos. Aficionado a los videojuegos y el tenis de mesa. Mi primer computador fue un Talent MSX que me compro mi papa por alla por el año 1985. En el di mis primeros pasos jugando juegos como Galaga y PacMan y luego programando en MSX-BASIC. 

En la actualidad mi area de conocimiento esta referida a las tecnologias .NET con mas de 15 años de experiencia desarrollando varias paginas web usando asp.net con bases de datos sql server y Oracle. Integrador de tecnologias, desarrollo de servicios, aplicaciones de escritorio.

# Tipo de Proyecto
Project Explorer es un analizador de código fuente para programas escritos en Visual Basic 4,5 y 6 desarrollado por mi por allá por el año 2000. El proyecto fue creado a modo de poder tener una herramienta personal que analizara el código fuente de mis proyectos escritos en Visual Basic y que realizara un análisis del código y pudiera generar un reporte de todas aquellas variables declaradas y que no estaban siendo ocupadas a modo de eliminarlas, limpiar el código y poder generar un ejecutable mas liviano. 

Luego expandi el análisis a todos los tipos de datos que soporta nativamente visual basic para poder analizar subrutinas, funciones, constantes, arreglos, enumeraciones, tipos, apis de windows y a la vez dependiendo el ámbito en el cual estaban declaradas (públicas, modulares y privadas) poder determinar si estas estaban siendo ocupadas. 

# Prologo
Regala un pescado a un hombre y le darás alimento para un día, enseñale a pescar y lo alimentarás para el resto de su vida (Proverbio Chino)

# Historia
Trabajaba por allá por el año 2000 en la compañia de seguros cruz del sur que estaba en el edificio el golf piso 10 metro Alcantara. En ese tiempo era programador externo migrando una aplicación de seguros desde una plataforma AS/400 a Visual Basic 6. Un día nos invitan a una charla de una empresa dedicada al QA (Quality Assurance) y nos muestran un software para hacer análisis de código fuente. 

La empresa nos dejo a varios de nosotros instalado el software a modo de poder realizar una revisión de proyectos reales y para posteriormente dar una opinión al respecto y ver la factibilidad técnica de poder comprarla.

En lo personal para mi era nuevo este tipo de software y no fue de mi gusto. El software era lentisimo en el análisis, se caia y con proyectos grandes no terminaba y cuando lograba terminar entregaba unos reportes que poco o nada se podia entender del análisis realizado. 

Finalmente terminaron por desecharlo y compraron otro software de origen finlandes el cual fue la base de comparación posteriormente para mi proyecto ...

# El inicio ...
Con la idea de fondo de poder generar una herramienta inicial para el análisis de código fuente y de limpiarlo de toda la basura (variables no usadas) y teniendo un punto de comparación con el cual ir revisando mis resultados fue que decidi escribir un analizador de aplicaciones para mi uso personal a modo de poder optimizar mis proyectos personales. 

El tema fue comenzar por leer el proyecto .vbp el cual tenia un formato de texto y era fácil de leer. Posterior a eso leer las rutas de los archivos relacionados (.frm, .bas. .cls, .ctl), el tipo de proyecto (.exe, .dll, .ocx) y luego por cada archivo leer el contenido (código fuente). 

Lo mejor de todo es que al ser archivos de texto plano no fue muy dificil poder implementar rutinas que pudieran "desarmar" el código fuente y poder determinar la estructura interna de este.

Luego una vez desarmado el proyecto lo demás fue generar clases que representaran de forma lógica la estructura para posteriormente poder extraer la información y poder procesarla. Como dato curioso fue un poco "loco" generar un proyecto para leer y analizar "proyectos" escritos en Visual Basic y analizarlos con una herramienta escrita en el mismo lenguaje.

# Algoritmo del analizador
Como la ideal inicial base de limpiar mi código fuente de variables no usadas para la construccion del analizador lo primero fue establecer la declaración de los ámbitos de las variables para su etapa inicial. Esto debido a que las variables se podian declarar de distintas formas y dependiendo de esto se asumian que eran públicos o privados dependiendo donde estaban declarados.

Una variable tiene 3 ámbitos para poder ser usada en un proyecto Visual Basic :

- Global : variable visible a todo el proyecto. (Generalmente declaradas en el módulo Global.bas)
- Modular : variable visible al ambito de donde esta declarada y de todos los elementos que existen en el. (Módulo)
- Privada : variable privada solo al ámbito donde esta declarada. (variables de subs y funciones)

Con esta idea base del ámbito de una variable luego fue el tema de realizar la búsqueda y si estaba siendo usada. Luego el analizador recorria las variables encontradas y dependiendo el "ambito" realizaba las búsquedas dentro del código fuente. Aca vino la pregunta el millón. Que es lo que realmente debo analizar ? 

Me explico, para realizar un análisis de solo lo necesario por ejemplo en el caso de buscar la variable en un procedimiento tenia que dejar afuera :

- Cabecera del procedimiento y la de término (sub y end sub)
- Declaración de parámetros usando el operador de multilinea ( _ )
- Lineas de comentarios
- Lineas donde habia una variable declarada

Ademas en el caso particular de los procedimientos y subs tenia que diferenciar por ejemplo que no hubiera una variable local o parámetro llamado de la misma manera ...

Con esta "limpieza" de lo realmente necesario recien podía buscar la variable y ver si estaba siendo usada. Para los mas entendidos Basic no se caracteriza por ser un lenguaje que sea óptimo en el recurso de memoria o que pueda ser "mas rápido de lo normal" en el procesamiento de strings por ejemplo. Para proyectos pequeños el analizador inicial funcionaba bien pero en proyectos "reales" o de muchas lineas de código era realmente lento.

Aca es donde la API de Windows y tipos "freak" por decirlo cariñosamente fueron de real ayuda a la hora de poder realizar analisis de variables de forma rápida.

Me tomo como 2 meses generar una versión simple y funcional de mi aplicación que me entregara un reporte con todo aquello que no estaba siendo usado y que estaba demas en mis proyectos. 

Luego con la base inicial del analizador lo fui mejorando para realizar análisis de arreglos, constantes, enumeraciones, apis de windows, parámetros,  tipos.

Toda esta segunda parte con los nuevos elementos y del como analizarlos me tomo como 1 año de implementar y generar un reporte con todo lo no usado en el proyecto.

# QA de Software
Resuelta la primera parte comenze a estudiar y investigar temas de QA (Quality Assurance) de software y en general para software desarrollado con Visual Basic. Tomando como base algunas ideas del software inicial y luego de mi experiencia propia es que el software ya no solo analizaba todas las variables sino que ademas establecer reglas y buenas prácticas para la escritura de código fuente.

La idea tras esta parte era por ejemplo :

- Nomenclatura de variables, objetos y controles de usuario.
- Máximo de lineas de código por procedimiento o función.
- Obligación de declaración de variables (Option Explicit)
- Comentarios de código 

Toda esta parte me tomo como 1 año de trabajo y de implementar y de ser casi 100% funcional. 

# Freeware
Por esos años mi intención fue ofrecerlo gratis a la comunidad Visual Basic que era bastante activa por esos años. Para esto levante un sitio web donde tenia varias otras aplicaciones que tambien habian sido creadas de la necesidad y que las distribuia de forma gratis.

# Palabras Finales
Espero que este proyecto que nacio de una necesidad personal sea usado con motivos de estudio y motivación. De como se pueden copiar las buenas ideas y mejorarlas. 
