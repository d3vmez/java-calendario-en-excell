package calendario;

import java.io.File;
import java.io.IOException;
import java.util.Calendar;

import org.apache.log4j.Logger;

import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Colour;
import jxl.format.UnderlineStyle;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

/**
 * 
 * @author Marcos
 *
 */

public class Calendario {

	private static Logger log = Logger.getLogger(Calendario.class);
	
	private static final String DIAS[] = { "L", "M", "X", "J", "V", "S", "D" };
	
	private static final String MESES[] = { "enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
			"septiembre", "octubre", "noviembre", "diciembre" };
	
	// Nombre del fichero
	private static final String NOMBRE_FICHERO = "Calendario.xls";
	
	private int año;
	private Calendar fecha = Calendar.getInstance();

	public Calendario(int año) {
		this.año = año;
	}

	/**
	 * 
	 * Función encargada de la escritura en Excell
	 */
	public void escribirEnExcell() {
		
		// creación del fichero
		File fichero = new File(NOMBRE_FICHERO);
		log.debug(fichero.getName() + " se ha creado correctamente");

		try {
			
			// creación del woorkbook 
			WritableWorkbook workbook = Workbook.createWorkbook(fichero);
			
			log.debug("el workbook se ha creado correctamente");	
			log.info("Fichero creado");
			
			// generar las 12 hojas para los meses en Excell		
			for (int i = 0; i < MESES.length; i++) {
				escribirHoja(workbook,i);
				log.debug("la hoja del mes de " + MESES[i] + " se ha creado correctamente");
			}
			
			workbook.write();
			log.info("Fin de escritura de las hojas");
			
			workbook.close();
			log.info("Fichero cerrado");

		} catch (IOException e) {
			
			log.error(e.getMessage());
			
		} catch (RowsExceededException e) {
			
			log.error(e.getMessage());
			
		} catch (WriteException e) {
			
			log.error(e.getMessage());
		}
	}

	/**
	 * 
	 * @param WritableWorkbook workbook, int numeroSheet
	 * Se le pasa la instancia del workbook y el indice para crear una nueva hoja
	 * @throws WriteException
	 * @throws RowsExceededException
	 */
	private void escribirHoja(WritableWorkbook workbook, int numeroSheet) throws RowsExceededException, WriteException {
		fecha.set(this.año, numeroSheet, 1);
		WritableSheet sheet = workbook.createSheet(MESES[numeroSheet], numeroSheet);
		int dias = 1;

		// formato para los días de la semana
		WritableFont fuenteDias = new WritableFont(WritableFont.TAHOMA, 10, WritableFont.BOLD, false,
				UnderlineStyle.NO_UNDERLINE, Colour.WHITE);
		WritableCellFormat formatoDias = new WritableCellFormat(fuenteDias);
		formatoDias.setAlignment(Alignment.CENTRE);
		formatoDias.setBackground(Colour.BLUE);

		// formato para los días festivos
		WritableFont fuenteFestivos = new WritableFont(WritableFont.TAHOMA, 10, WritableFont.BOLD, false,
				UnderlineStyle.NO_UNDERLINE, Colour.GREEN);
		WritableCellFormat formatoFestivos = new WritableCellFormat(fuenteFestivos);
		formatoFestivos.setAlignment(Alignment.CENTRE);
		
		// formato para alinear al centro todos los dias
		WritableCellFormat formatoCentrado = new WritableCellFormat();
		formatoCentrado.setAlignment(Alignment.CENTRE);
	
		// escribir mes y año
		Label mes = new Label(0, 0, MESES[numeroSheet]);
		sheet.addCell(mes);
		
		Label año = new Label(1, 0, String.valueOf(this.año));
		sheet.addCell(año);
		
		// escritura de los nombres de los dias
		for (int j = 0; j < DIAS.length; j++) {
			Label etiqueta = new Label(j, 1, DIAS[j]);
			etiqueta.setCellFormat(formatoDias);
			sheet.addCell(etiqueta);
		}

		// escritura de la primera semana del mes
		for (int j = calcularDiaEmpiezaMes(numeroSheet); j < DIAS.length; j++) {
			Number number = new Number(j, 2, dias);
			
			// Si es sábado o domingo se pone en negrita verde
			if (j == 5 || j == 6) {
				number.setCellFormat(formatoFestivos);
				sheet.addCell(number);
				dias++;
			} else {
				number.setCellFormat(formatoCentrado);
				sheet.addCell(number);
				dias++;
			}
		}

		// escritura de las restantes semanas
		// le sumo 1 al numero de semanas porque Calendar como máximo se saca 5 semanas
		for (int i = 3; i <= calcularSemanasDelMes(numeroSheet) + 1; i++) {
			for (int j = 0; j < DIAS.length; j++) {

				if (!(dias > calcularDiasDelMes(numeroSheet))) {
					
					Number number = new Number(j, i, dias);

					if (j == 5 || j == 6) {
						number.setCellFormat(formatoFestivos);
						sheet.addCell(number);
						dias++;
					} else {
						number.setCellFormat(formatoCentrado);
						sheet.addCell(number);
						dias++;
					}
				}
			}
		}

	}

	/**
	 * 
	 * @param int mes, en el que estamos
	 * @return el numero de semanas del mes
	 */
	private int calcularSemanasDelMes(int mes) {

		fecha.setFirstDayOfWeek(calcularDiaEmpiezaMes(mes));

		int semanas = fecha.getActualMaximum(Calendar.WEEK_OF_MONTH);
		log.debug("total de semanas:"+semanas);
		return semanas;
	}
	
	/**
	 * 
	 * @param  int mes, en el que estamos
	 * @return los dias que tiene el mes
	 */
	private int calcularDiasDelMes(int mes) {
		
		int dias = fecha.getActualMaximum(Calendar.DAY_OF_MONTH);
		
		log.debug("total dias del mes: "+dias);
		return dias;
		
	}
	
	/**
	 * 
	 * @param  int mes, en el que estamos
	 * @return el dia en el que empieza el mes
	 */
	private int calcularDiaEmpiezaMes(int mes) {

		int dia = fecha.get(Calendar.DAY_OF_WEEK);

		// conversión del formato americano a mi formato, se empieza en lunes (0)

		// domingo
		if (dia == 1) {
			dia = 6;

		} else

		// lunes
		if (dia == 2) {
			dia = 0;
		} else

		// martes
		if (dia == 3) {
			dia = 1;
		} else

		// miercoles
		if (dia == 4) {
			dia = 2;
		} else

		// jueves
		if (dia == 5) {
			dia = 3;
		} else

		// viernes
		if (dia == 6) {
			dia = 4;
		} else

		if (dia == 7) {
			dia = 5;
		}
		
		log.debug("la semana empieza: " + dia);

		return dia;
	}

}
