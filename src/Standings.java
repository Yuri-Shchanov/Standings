

import java.io.File;
import java.io.IOException;
import java.lang.String;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
//import jxl.write.Formula;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;
import jxl.format.Border;
import jxl.format.BorderLineStyle;



public class Standings{
	private boolean Import = false; 
	private String fileAddress;
	private Integer number;
	private Integer col[];
	private Integer row[];
	private String score[]; 
	private String player[];
	
	public Integer getNumber() {
		return number;
	}

	public void setNumber(Integer number) {
		this.number = number;
	}
		
	private void Captions(int number) throws BiffException, IOException, RowsExceededException, WriteException{
		col = new Integer[number * 2];
		row = new Integer[number * 2];
		player = new String[number * 2];
		score = new String[number * 2];
		Workbook wrk1 = null;
		Sheet sheet1 = null;
		if (Import){
			wrk1 =  Workbook.getWorkbook(new File(fileAddress));
			sheet1 = wrk1.getSheet(0);
		}
		int count = number;
	  	int col1 = 2;
		int row1 = 0;
		int rowch = 2;
		for (int i = 1; i< 2 * number; i++){
			player[i] = sheet1.getCell(col1, row1).getContents();
			score[i] = (i > number)? " " + sheet1.getCell(col1+1, row1+2).getContents() : "";
            row[i] = row1;
            col[i] = col1;
            row1 += rowch;
			if (i == count){
				col1 += (i == number) ? 3 : 4;
				count += (number * 2 - count) / 2;
				row1 = rowch - 1;
				rowch *= 2;
			}
			
			
		}
		if(Import){
			wrk1.close();
		}
		Import = false;
	}
	
	void insert(int slot, String name){
		player[slot] = name;
	}
	
	void setScore(int slot1, int slot2, String score){
		int slot = slot2 / 2 + 128;
		int score1 = Integer.parseInt(score.split(":", 2)[0]);
		int score2 = Integer.parseInt(score.split(":", 2)[1]);
		player[slot] = (score1 > score2)? player[slot1] : player[slot2];
		this.score[slot] = "—чет матча: " + score;
	}
		
	Standings(int number) throws BiffException, IOException, RowsExceededException, WriteException{
	    setNumber(number);
		Captions(number);
  	}

    Standings(String fileAddress) throws BiffException, IOException, RowsExceededException, WriteException {
		this.fileAddress = fileAddress;
    	Import = true;
    	
    	number = Integer.parseInt(fileAddress.split("/",4)[3].split(".xls", 2)[0]);
    	Captions(number);
	}
    
  	public void export(String fileAddress) {
        try {
        	WritableWorkbook workbook = Workbook.createWorkbook(new File(fileAddress), 
        			Workbook.getWorkbook(new File("c:/Java projects/Standings/" + number + ".xls")));
        	WritableSheet sheet = workbook.getSheet(0);
        	WritableFont font =
        			new WritableFont(WritableFont.ARIAL, 10, WritableFont.NO_BOLD);
        	WritableCellFormat format = new WritableCellFormat(font);
        	format.setBorder(Border.LEFT, BorderLineStyle.HAIR);
        	format.setBorder(Border.BOTTOM, BorderLineStyle.HAIR);
            for(int i = 1; i < number * 2; i++){
           	if(i > number){
           		sheet.addCell(new Label(col[i], row[i], player[i]));
            	sheet.addCell(new Number(col[i], row[i] + 2, i));
            	sheet.addCell(new Label(col[i] + 1, row[i] + 2, score[i]));
            }
            else {
            	sheet.addCell(new Number(col[i] - 2, row[i], i));
            	sheet.addCell(new Label(col[i], row[i], player[i], format));
            }
       }      
            workbook.write();
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (RowsExceededException e) {
            e.printStackTrace();
        } catch (WriteException e) {
            e.printStackTrace();
        } catch (BiffException e) {
        	e.printStackTrace();
        }
    }
}
  