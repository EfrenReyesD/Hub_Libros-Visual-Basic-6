# Hub_Libros-Visual-Basic-6
## Descripción
El proyecto es una aplicación para gestionar y visualizar libros, permitiendo a los usuarios interactuar con una base de datos para realizar diversas acciones relacionadas con la lectura. Los usuarios pueden ver un catálogo de libros, marcar libros como leídos, favoritos, o no deseados
## Objetivo 
1.- Aprender la tecnologia visual basic 6
2.- La aplicación se realiza de tal manera que permita a los usuarios gestionar su colección de libros, marcandolos como leidos, favoritos o libros que no les gustan
3.- Proporcionar una interfaz intuitiva, con la creacion de menus y con la informacion organizada es de facil interazción con un catalogo de libros donde tambien el usuario puede buscar a traves de la barra de busqueda por si le interesa algo muy especifico
## Funcionamiento
 Al final se agrega un video de la interfaz y la interaccion que puede tener el usuario
 
![image](https://github.com/user-attachments/assets/1c2c716b-c350-4909-968c-5fc4a99d2f55)

![image](https://github.com/user-attachments/assets/a46c3937-ade8-489e-98b6-92cba4379f3c)


## Instrucciones de uso
1.- clonar el repositorio
2.- Asegurate de tener instalado las configuraciones de entorno como lo son: visual basic 6 y SQL server
3.- crea la base da datos, en los archivos de la carpeta Modules puedes realizar las configuraciones necesarias para la DB
## Descripción de elaboración
El proyecto es una aplicación de gestión y visualización de libros, desarrollada en Visual Basic 6 y respaldada por una base de datos en SQL Server. La aplicación permite a los usuarios gestionar sus preferencias de lectura y visualizar información sobre libros.
2.- Diseño de la base de datos
Se diseñó el esquema de la base de datos utilizando un modelo entidad-relación (ERD), que incluye tres tablas principales:
Users:
Almacena información del usuario, como nombre, apellido, contraseña, y URL de la foto de perfil.
Se añadió una columna para la fecha de creación del usuario.
Books:
Contiene información sobre los libros, incluyendo título, autor, género, URL del PDF, imagen de portada, y descripción.
UserBooks:
Maneja las preferencias del usuario sobre los libros, permitiendo marcar un libro como leído, favorito, o no deseado.
#### Desarrollo de la Aplicacion
Se crearon varios formularios para diferentes funcionalidades, como el catálogo de libros, gestión de preferencias, y autenticación de usuarios.
Controles: Se utilizaron controles como DataGrid para mostrar datos en formato tabular, TextBox para entradas de usuario, y CommandButton para acciones como agregar a favoritos.
#### Validacion
Se realizaron pruebas para verificar la funcionalidad de la aplicación, incluyendo la gestión de usuarios, la visualización de libros, y la actualización de preferencias.
Verificación de Datos:
Se verificó que los datos en la base de datos se reflejen correctamente en la interfaz de usuario y que las consultas SQL devuelvan los resultados esperados.
#### Compartido 
El proyecto se subió a un repositorio de Git para facilitar el acceso y la colaboración. Se proporcionaron detalles sobre la base de datos y las instrucciones de uso en este archivo README.
## Diagrama de DB Entidad-Relación
![image](https://github.com/user-attachments/assets/830bb869-032b-479e-b9a9-528e34d4e38c)

## Problemas conocidos
- Cada vez se va dando uno cuenta la importancia que debe haber al realizar un proyecto, se requiere de una buena planeacion ya que si no se realiza de esa manera se realizaran diferentes cambios que al final suele complicarse la estructuración.
- Fue un gran reto realizar el proyecto ya que con nuevos lenguajes es un poco dificil agarrarle el rollo.
- uno de los problemas que se tuvo al final fue que al tratar de eliminar un usuario supuestamente se complicaba realizar la tarea por temas de llaves foraneas. ya no dio tiempo de subir la solucion a esto.

## Retrospectiva

### ¿Qué hice bien?
- Comparante los anteriores Sprint debo decir que ahora tuve un poco mas de organizacion o de primero pensar que herramientas utilizar para no complicarme tanto la estruccturación del proyecto

### ¿Qué no salio bien?
- Despues de varias pruebas al final intente eliminar un usuario y ya no me permitio, segun el error era directamente de las base de datos y tema de claves foraneas por las relaciones que habia con otras tablas, faltaron algunas funcionalidades que en su estan implementadas en otras tarea por ejemplo en el registro de usuarios quedo y falto el registro de un nuevo libro

### ¿Qué puedo hacer diferente?
- Creo que voy por buen camino en cuanto a organizarme. tambien creo que debo reforzar las diferentes formas de conexion hacia base de datos y con que tenologias puedo combinarlas, tambien puedo ser mas facil de otra manera, queda investigar

### Evidencia 



https://github.com/user-attachments/assets/cb2f7bb0-b977-47b6-9e9a-dc49c1c5043e


