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
		int a�o;
		
		System.out.print("Introduce el a�o para generar el calendario en Excell: ");
		a�o = sc.nextInt();
		
		Calendario calendario = new Calendario(a�o);
		calendario.escribirEnExcell();
		
		
		sc.close();
	}

}
