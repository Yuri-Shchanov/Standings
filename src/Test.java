import java.io.IOException;

import jxl.read.biff.BiffException;
import jxl.write.WriteException;


public class Test {

	 public static void main(String[] args) throws WriteException, IOException, BiffException {
//		  Standings test = new Standings(128);
		  Standings test = new Standings("128.xls");
		  test.insert(3, "sфывда");
		  test.insert(4, "sdfgss");
		  test.setScore(3, 4, "2432:123453");
		  test.export("128.xls");
		  //test.write();
	  }
}
