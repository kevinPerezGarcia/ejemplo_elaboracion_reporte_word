/*******************************************************************************
* Fecha:		2021-10-23
* Instituto:	Grupo CEA
* Curso:		Programación estadística (con Stata)
* Sesión:		Ejemplo de elaboración de reporte en Word
* Autor:		Kevin Pérez García
* email:		econ.perez.garcia.k@gmail.com
********************************************************************************

*** Objetivo:
	Construir variables
	
*** Requerimientos:
	inputs: 	"${data_construccion}/sumaria-m100.dta"
	outputs: 	"${data_tablas}/sumaria-m100-construida.dta"
	
*** Outline:
	1.1 Cargando base de datos
	1.2 Iniciando documento word
	1.3 Insertando título
	1.4 Insertando sección con cita
		1.4.1 Insertando encabezado (subtítulo)
		1.4.2 Insertando párrafo con cita
	1.5 Insertando sección con estadísticos
		1.5.1 Insertando encabezado (subtítulo)
		1.5.2 Calculando estadísticos
		1.5.3 Insertando párrafo con estadísticos
	1.6 Insertando sección con gráfico
		1.6.1 Insertando encabezado (subtítulo)
		1.6.2 Graficando
		1.6.3 Exportando gráfico a uno admitido en word
		1.6.4 Estableciendo alineación central
		1.6.5 Insertando gráfico
	1.7 Insertando tabla de estimación
		1.7.1 Insertando encabezado (subtítulo)
		1.7.2 Estimando regresión
		1.7.3 Insertando tabla de estimación
	1.8 Apilando cambios realizados en el documento

*** Observaciones:

	Mayor personalización:

		Consulte "putdocx begin" para obtener información sobre cómo formatear el documento en su conjunto, incluida la especificación del tamaño de página, el diseño de la página, la fuente y los encabezados y pies de página.

		Consulte "putdocx paragraph" para obtener información sobre cómo agregar bloques completos de texto a un documento; modificando el estilo, fuente, alineación y otro formato de un párrafo; personalizar el tamaño y la ubicación de un imagen; y agregar contenido a un encabezado o pie de página. 

		Consulte la "putdocx table" para obtener información sobre cómo crear tablas de resultados, matrices, datos e incluso imágenes y para obtener información sobre cómo personalizarlos. 

		Consulte "putdocx pagebreak" para obtener información sobre cómo agregar saltos de página y saltos de sección a su documento.

********************************************************************************
***	PART 1: Construyendo variables
*******************************************************************************/
	
*** 1.1 Cargando base de datos
	use "http://www.stata-press.com/data/r15/lbw.dta", clear

*** 1.2 Iniciando documento word
	putdocx begin
	putdocx save "${outputs_reportes}/reporte.docx", replace // requiere empezar nuevamente con putdocx begin

*** 1.3 Insertando título
	putdocx begin
	putdocx paragraph, style(Title)
	putdocx text ("Informe sobre el bajo peso al nacer")

*** 1.4 Insertando sección con cita

	*** 1.4.1 Insertando encabezado (subtítulo)
		putdocx paragraph, style(Heading1)
		putdocx text ("Introducción a los datos")

	*** 1.4.2 Insertando párrafo con cita
		putdocx paragraph
		putdocx text ("Tenemos datos sobre el peso al nacer de Hosmer, Lemeshow, and ")
		putdocx text ("Sturdivant (2013, 24).")

*** 1.5 Insertando sección con estadísticos

	*** 1.5.1 Insertando encabezado (subtítulo)
		putdocx paragraph, style(Heading1)
		putdocx text ("Resumen de estadísticas")

	*** 1.5.2 Calculando estadísticos
		summarize bwt
		return list

	*** 1.5.3 Insertando párrafo con estadísticos
		putdocx paragraph
		putdocx text ("Tenemos el peso registrado de `r(N)' bebés ")
		putdocx text ("con un peso medio al nacer de ")
		putdocx text (" `r(mean)' "), nformat(%5.2f)
		putdocx text (" gramos .")

*** 1.6 Insertando sección con gráfico
	
	*** 1.6.1 Insertando encabezado (subtítulo)
		putdocx paragraph, style(Heading1)
		putdocx text ("Peso al nacer por estado de tabaquismo de la madre")

	*** 1.6.2 Graficando
		graph hbar bwt, over(ht,relabel(1 "No hypertension" 2 "Has history of hypertension")) over(smoke) asyvars ytitle(Average birthweight (grams)) title(Baby birthweights)   subtitle(by mother's smoking status and history of hypertension)
		
		/* Nota.
		Graficamos el peso medio al nacer de los bebés cuyas madres fuman frente
		a las que no lo hacen, y por separado para madres con y sin antecedentes
		de hipertensión.
		*/

	*** 1.6.3 Exportando gráfico a uno admitido en word
		quietly graph export "${outputs_graficos}/bweight.png", replace
		
		* Nota. Formatos de imagen admitidos: .jpg, .emf, .tif o .png.
		
	*** 1.6.4 Estableciendo alineación central
		putdocx paragraph, halign(center)
	
	*** 1.6.5 Insertando gráfico
		putdocx image bweight.png, width(4) height(2.8)
		
		* Nota. Estableciendo el ancho de la imagen en 4 pulgadas y el altura en 2,8 pulgadas.

*** 1.7 Insertando tabla de estimación
	
	*** 1.7.1 Insertando encabezado (subtítulo)
		putdocx paragraph, style(Heading1)
		putdocx text ("Regression results")

	*** 1.7.2 Estimando regresión
		regress bwt smoke age, noheader
		
		/* Nota.
		Regresión después de modelar el peso al nacer (bwt) como una función de 
		la edad de la madre (age) y si fuma (smoke).
		*/

	*** 1.7.3 Insertando tabla de estimación
		putdocx table bweight = etable, title("Regresión lineal del peso al nacer")
		
		* Nota. Usamos la opción <title()> para agregar un título a nuestra tabla.

*** 1.8 Apilando cambios realizados en el documento
	putdocx save "${outputs_reportes}/reporte.docx", append
