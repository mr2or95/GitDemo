package MavenTest.ExcelData;

import java.io.IOException;
import java.util.ArrayList;

public class TestCaseData {

	public static void main(String[] args) throws IOException {

		// Access the data from a file
		AccessData atc = new AccessData();
		
		ArrayList<String> tcset = atc.getData("TC6");
		
		// make sure the array is not empty
		if (!tcset.isEmpty()) {
			//print out the data for the testcase
			for (int i = 0; i < tcset.size(); i++) {
				System.out.println(tcset.get(i));
			}
		} else {
			System.out.println("There is no such testcase\n");
		}
	}

}
