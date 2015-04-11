import java.io.IOException;

import jxl.read.biff.BiffException;
import jxl.write.WriteException;


public class Test {

	 public static void main(String[] args) throws WriteException, IOException, BiffException {
//		  Standings test = new Standings(128);
		  Standings test = new Standings("C:/Java projects/Standings/128.xls");
		  test.insert(1, "sd52");
		  test.insert(2, "s");
		  test.setScore(1, 2, "24:453");
		  test.export("C:/Java projects/Standings/128.xls");
		  //test.write();
	  }
}
