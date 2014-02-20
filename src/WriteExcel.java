import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Locale;
import java.util.Map;

import jxl.CellView;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.UnderlineStyle;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class WriteExcel implements Constants {


	private WritableCellFormat timesBoldUnderline;
	private WritableCellFormat times;
	private String inputFile;
	private String missingFile;
	private BudgetAnalysis analysis; 
	
	public WriteExcel(BudgetAnalysis analysis) {
		this.analysis = analysis;
	}
	
	public void setOutputFiles(String inputFile, String missingFile) {
		this.inputFile = inputFile;
		this.missingFile = missingFile;
	}

	public void write() throws IOException, WriteException {
		
		File file = new File(inputFile);
		WorkbookSettings wbSettings = new WorkbookSettings();

		wbSettings.setLocale(new Locale("en", "EN"));

		WritableWorkbook workbook = Workbook.createWorkbook(file, wbSettings);
		workbook.createSheet("Projections", 0);
		WritableSheet excelSheet = workbook.getSheet(0);
		
		createLabels(excelSheet);
		
		createContent(excelSheet);		

		workbook.write();
		
		workbook.close();
	}

	public void writeMissingData() throws IOException, WriteException {

		File file = new File(missingFile);
		WorkbookSettings wbSettings = new WorkbookSettings();
		wbSettings.setLocale(new Locale("en", "EN"));

		WritableWorkbook workbook = Workbook.createWorkbook(file, wbSettings);
		workbook.createSheet("Missing Accounts in Projection", 0);
		WritableSheet excelSheet = workbook.getSheet(0);
		
		createMissingContent(excelSheet);		
		workbook.write();		
		workbook.close();
	}
	
	public void createMissingContent(WritableSheet sheet) throws WriteException {
		
		WritableFont times12pt = new WritableFont(WritableFont.TIMES, 12);
		times = new WritableCellFormat(times12pt);
		times.setWrap(false);

		// Create create a bold font with unterlines
		WritableFont times12ptBoldUnderline = new WritableFont(
				WritableFont.TIMES, 12, WritableFont.BOLD, false,
				UnderlineStyle.SINGLE);
		timesBoldUnderline = new WritableCellFormat(times12ptBoldUnderline);
		timesBoldUnderline.setWrap(false);

		CellView cv = new CellView();
		cv.setFormat(times);
		cv.setFormat(timesBoldUnderline);
		cv.setAutosize(true);

		// Write a few headers
		addCaption(sheet, 0, 1, "Warehouse");
		addCaption(sheet, 1, 1, "Acct#Name");

		ArrayList<String[]> missingList = analysis.getMissingAccountList();
		Iterator<String[]> it = missingList.iterator();
	    
		int indexY = 2;
		while (it.hasNext()) {
	    	String[] account = (String[])it.next();	    	
	    	String _warehouseName = account[0];
	    	String _accountName = account[1];
			
	    	addCaption(sheet, 0, indexY, _warehouseName);	    	
	      	addCaption(sheet, 1, indexY, _accountName);	    	
	  	    	
	      	indexY++;
	    }		
	}
	
	public void createLabels(WritableSheet sheet) throws WriteException {
		
		// Lets create a times font
		WritableFont times12pt = new WritableFont(WritableFont.TIMES, 12);
		// Define the cell format
		times = new WritableCellFormat(times12pt);
		
		// Lets automatically wrap the cells
		times.setWrap(false);

		// Create create a bold font with underlines
		WritableFont times10ptBoldUnderline = new WritableFont(
				WritableFont.TIMES, 12, WritableFont.BOLD, false,
				UnderlineStyle.SINGLE);
		timesBoldUnderline = new WritableCellFormat(times10ptBoldUnderline);
		
		// Lets automatically wrap the cells
		timesBoldUnderline.setWrap(false);

		CellView cv = new CellView();
		cv.setFormat(times);
		cv.setFormat(timesBoldUnderline);
		cv.setAutosize(true);

		// Write a few headers
		addCaption(sheet, 0, 1, "Warehouse");
		addCaption(sheet, 1, 1, "Acct#Name");
		
		addCaption(sheet, 2, 0, "Volumes");
		
		// Add the rest 
		String[] volumes = analysis.getLabels(analysis.getHistSheet(), strVOLUME_LABEL_RANGE);
		for (int i=0; i<volumes.length; i++) 
			addCaption(sheet, i+2, 1, volumes[i]);

		addCaption(sheet, 2+volumes.length, 0, "Gross Margins");
		
		String[] margins = analysis.getLabels(analysis.getHistSheet(), strGROSS_MARGIN_LABEL_RANGE);
		for (int j=0; j<margins.length; j++) 
			addCaption(sheet, j+2+volumes.length, 1, margins[j]);

	}

	public void createContent(WritableSheet sheet) throws WriteException, RowsExceededException {
		
		Map<String, Map<String, ArrayList<double[]>>> warehouseMap = analysis.getWarehouseMap();
		
		int y_index = 2;
		for (Map.Entry<String, Map<String, ArrayList<double[]>>> masterEntry : warehouseMap.entrySet()) {			
			
			String warehouseName = masterEntry.getKey();
			
			addLabel(sheet, 0, y_index, warehouseName);
			
			Map<String, ArrayList<double[]>> accountMap = masterEntry.getValue();

			for (Map.Entry<String, ArrayList<double[]>> accountEntry : accountMap.entrySet()) {
				
				String accountName = accountEntry.getKey();				
				addLabel(sheet, 1, y_index, accountName);
				
				ArrayList<double[]> list = accountEntry.getValue();
								
				// print volumes
				double[] volumes = list.get(0);
				for (int i=0; i<volumes.length; i++) {
					addNumber(sheet, i+2, y_index, volumes[i]);				
				}
				
				// print gross margins
				double[] margins = list.get(1);
				for (int i=0; i<margins.length; i++) {
					addNumber(sheet, i+2+volumes.length, y_index, margins[i]);				
				}

				y_index++;
			}	
			
		}
		
	}

	private void addCaption(WritableSheet sheet, int column, int row, String s)
			throws RowsExceededException, WriteException {
		Label label;
		label = new Label(column, row, s, timesBoldUnderline);
		sheet.addCell(label);
	}

	private void addNumber(WritableSheet sheet, int column, int row,
			double d) throws WriteException, RowsExceededException {
		Number number;
		number = new Number(column, row, d, times);
		sheet.addCell(number);
	}

	private void addLabel(WritableSheet sheet, int column, int row, String s)
			throws WriteException, RowsExceededException {
		Label label;
		label = new Label(column, row, s, times);
		sheet.addCell(label);
	}
}