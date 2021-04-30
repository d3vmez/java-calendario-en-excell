package calendario;

import java.io.File;
import java.io.IOException;
import java.time.DayOfWeek;
import java.util.Calendar;

import org.apache.log4j.Logger;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class Calendario {

	private static Logger log = Logger.getLogger(Calendario.class);
	
	private static final String DIAS[] = { "L", "M", "X", "J", "V", "S", "D" };
	
	private static final String MESES[] = { "enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
			"septiembre", "octubre", "noviembre", "diciembre" };
	
	private int año;

	public Calendario(int año) {
		this.año = año;
		escribirEnExcell();
		log.debug("so");
	}

	public int getAño() {
		return año;
	}

	public void setAño(int año) {
		this.año = año;
	}

	/**
	 * 
	 */
	public void escribirEnExcell() {

		File fichero = new File("Calendario.xls");
		log.debug(fichero.getName() + " se ha creado correctamente");

		try {
			WritableWorkbook workbook = Workbook.createWorkbook(fichero);
			log.debug("el workbook se ha creado correctamente");

			for (int i = 0; i < MESES.length; i++) {
				escribirHoja(workbook,i);
			}
			
			workbook.write();
			workbook.close();

		} catch (IOException e) {
			e.printStackTrace();
		} catch (RowsExceededException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	/**
	 * 
	 * @param workbook
	 * @throws WriteException
	 * @throws RowsExceededException
	 */
	public void escribirHoja(WritableWorkbook workbook, int numeroSheet) throws RowsExceededException, WriteException {

		WritableSheet sheet = workbook.createSheet(MESES[numeroSheet], numeroSheet);
		int dias = 1;
		
		// se escribe en la fila el nombre de los dias
		for (int j = 0; j < DIAS.length; j++) {
			Label etiqueta = new Label(j, 0, DIAS[j]);
			sheet.addCell(etiqueta);
		}
		
		// escritura de los dias del mes
		for (int i = 1; i <= calcularSemanasDelMes(numeroSheet); i++) {
			for (int j = 0; j < DIAS.length; j++) {
				
				if(! (dias > calcularDiasDelMes(numeroSheet))) {
				Number number = new Number(j,i,dias);
				sheet.addCell(number);
				dias++;
				
				}
			}
			
		}

	}

	/**
	 * 
	 * @return Devuelve el número de semanas que tiene un mes
	 */
	public int calcularSemanasDelMes(int mes) {

		Calendar fecha = Calendar.getInstance();
		fecha.set(this.año, mes, 1);
		fecha.setFirstDayOfWeek(calcularDiaEmpiezaMes(mes));

		int semanas = fecha.getActualMaximum(Calendar.WEEK_OF_MONTH);
		log.debug("total de semanas:"+semanas);
		return semanas;
	}
	
	/**
	 * 
	 * @param mes
	 * @return
	 */
	public int calcularDiasDelMes(int mes) {
		Calendar fecha = Calendar.getInstance();
		fecha.set(this.año, mes, 1);

		int dias = fecha.getActualMaximum(Calendar.DAY_OF_MONTH);
		log.debug(dias);
		return dias;
		
	}
	
	public int calcularDiaEmpiezaMes(int mes) {
		Calendar fecha = Calendar.getInstance();
		fecha.set(this.año, mes, 1);
		
		int dia = fecha.get(Calendar.DAY_OF_WEEK);
		log.debug("empieza en "+ dia);
		
		return dia;
	}

}
