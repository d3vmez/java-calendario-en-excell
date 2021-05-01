package calendario;

import java.util.Scanner;

/**
 * 
 * @author Marcos
 *
 */
public class CalendarioApp {

	public static void main(String[] args) {
		
		// declaraciones
		Scanner sc = new Scanner(System.in);
		int año;
		
		System.out.print("Introduce el año para generar el calendario en Excell: ");
		año = sc.nextInt();
		
		Calendario calendario = new Calendario(año);
		calendario.escribirEnExcell();
		
		
		sc.close();
	}

}
