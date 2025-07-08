
## Index
* [1. Motivation / Motivación](#Motivation-/-Motivación)
* [2. What you do need? / ¿Qué se necesita?](#what-you-do-need?-/-¿Qué-se-necesita?)
* [3. Use / Uso](#Use-/-Uso)
# Motivation / Motivación

[English]

In the current market of solutions to transform DOC or DOCX extension documents to PDF there is an infinite number of subscription-based options. Looking for very tight solutions you can find powerful libraries with a free layer and others with absurdly expensive licenses for those who write these words. I can't deny that sometimes you need practical solutions instead of having to hire a tank for what a bicycle needs. That is why I share the following code, simple in its performance, perhaps not as efficient as the libraries that can be found, but simple in its understanding.

[Español]

En el mercado actual de soluciones para transformar documentos de extensión DOC o DOCX a PDF existe una infinidad de opciones basadas en suscripción. Buscando soluciones muy ajustadas podrás encontrar potentes librerías con una capa gratuita y otras tantas con licencias absurdamente caras para quien escribe estas palabras. No puedo negar que debes en cuando se requieren soluciones practicas en vez de tener que contratar un tanque para aquello que necesita una bicicleta. Es por ello que comparto el siguiente código, simple en su actuar, quizás no tan eficiente como las librerías que se pueden encontrar, pero simple en su entendimiento.


## What you do need? / ¿Qué se necesita?

* C## (.net 8)
* [LibreOffice Portable](https://www.libreoffice.org/download/portable-versions/)

[English]

I know that this solution is totally free and this project will be in the public domain, but if this solution meets a need you have it would be nice if you could make a donation to [LibreOffice](https://www.libreoffice.org/donate/) to say thank you and allow to continue growing this Open Source software that has allowed the development of this solution. Donate whatever you feel is right.

[Español]

Se que esta solución es totalmente gratuita y este proyecto será de dominio publico, pero si esta solución satisface una necesidad que tengas sería de buen agrado que puedas hacer una donación a [LibreOffice](https://www.libreoffice.org/donate/) para agradecer y permitir que siga creciendo este software Open Source que ha permitido el desarrollo de esta solución. Dona lo que tu sientas que es correcto.
## Use / Uso

[English]

Its use is quite simple and it is commented the lines that you need to keep an eye on.

* DocxPath: Here you must replace by the access path where your DOCX is located, it is important to highlight that it must end in “\{Archivo_DOC.docx}”.

* OutPutFolder: Here you must replace it with the path where you want to deposit the PDF that will be generated.

* libreOfficePath: Here it must be replaced by the path where soffice.exe is located, generally it is where the LibreOfficePortable folder is and the path from there is “LibreOfficePortableApp ”LibreOfficePortable.exe".

* Arguments: Here you do not have to change anything, this is a line that you can throw by CMD from the folder where soffice.exe is located, it is advisable to test it before to make sure that soffice.exe is responding correctly, for this you have to open a CMD, navigate to the address of soffice typing “cd {path_soffice.exe}” and then copy what Arguments says replacing the parameters.

* Bonus: If you want to change the PDF output name, you can add a parameter to the ConvertDocxToPdf function and change the name it assigns to “fileRename”.

[Español]

Su uso es bastante sencillo y esta comentado las líneas que se necesita tener ojo.

* DocxPath: Aquí debes sustituir por la ruta de acceso en donde se encuentra tu DOCX, es importante resaltar que debe terminar en "\{Archivo_DOC.docx}"

* OutPutFolder: Aquí se debe sustituir por la ruta de acceso que se desea depositar el PDF que se generara.

* libreOfficePath: Aquí se debe sustituir por la ruta de acceso donde se encuentra soffice.exe, generalmente es donde esta la carpeta LibreOfficePortable y la ruta desde ahí es "LibreOfficePortable\App\libreoffice\program\soffice.exe"

* Arguments: Aquí no hay que cambiar nada, esta es una linea que se puede tirar por CMD desde la carpeta donde se encuentra soffice.exe, es recomendable probarla antes para asegurar que soffice.exe esta respondiendo correctamente, para ello hay que abrir una CMD, navegar hasta la dirección de soffice escribiendo "cd {ruta_soffice.exe}" y luego copiar lo que dice Arguments reemplazando los parametros.

* Bonus: Si deseas cambiar el nombre de salida del PDF, puedes agregar un parametro a la función ConvertDocxToPdf y cambiar el nombre que asigna en "archivoRenombrado" 
