import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.DataInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.WriteException;


public class BudgetAnalysis implements Constants {

	enum account_type {
		HISTORY_ACCOUNTS,
		PROJECTION_ACCOUNTS
	}
	
	enum warehouse_type {
		TO_PERCENTAGE_WAREHOUSE,
		TO_PROJECTION_WAREHOUSE
	}

	private String historyInputFile;
	private String projectionInputFile;
	
	private Sheet histSheet = null;
	private Sheet projSheet = null;
	
	// Map: <warehouseName, Map<accountName, volumes/margins>> 
	private static Map<String, Map<String, ArrayList<double[]>>> gWarehouseMap = new HashMap<String, Map<String, ArrayList<double[]>>>();
	
	public Map<String, Map<String, ArrayList<double[]>>> getWarehouseMap() {
		return gWarehouseMap;
	}


	private static Map<String, ArrayList<double[]>> gAccountMap = new HashMap<String, ArrayList<double[]>>();
	private static Map<String, ArrayList<double[]>> gProjectedAccountMap = new HashMap<String, ArrayList<double[]>>();
	
	private static ArrayList<String[]> missingAccountList = new ArrayList<String[]>();
	
	public ArrayList<String[]> getMissingAccountList() {
		return missingAccountList;
	}


	private static final String configDataPath = "/Users/joe/Documents/workspace/BudgetAnalysis/bin/config.txt";
	private static final String outputDataPath = "/Users/joe/Documents/workspace/BudgetAnalysis/bin/PROJECTED_ACCOUNTS.xls";
	private static final String missingDataPath = "/Users/joe/Documents/workspace/BudgetAnalysis/bin/MISSING_ACCOUNTS.xls";
	
	private static final String warehousesPath = "/Users/joe/Documents/workspace/BudgetAnalysis/bin/warehouses.txt";
	private static final String histAccountPath = "/Users/joe/Documents/workspace/BudgetAnalysis/bin/histAccounts.txt";
	private static final String projAccountPath = "/Users/joe/Documents/workspace/BudgetAnalysis/bin/projAccounts.txt";
	
	private static Map<String, String> configHash = new HashMap<String, String>();
		
	public void setInputFiles(String histInputFile, String projInputFile) {
		
		System.out.println();
		
		this.historyInputFile = histInputFile;
		this.projectionInputFile = projInputFile;
	}

	public void read() throws IOException  {
		
		File historyInputWorkbook = new File("./bin/" + historyInputFile);
		File projectionInputWorkbook = new File("./bin/" + projectionInputFile);
		
		Workbook histW, projW;
				
		try {
						
			histW = Workbook.getWorkbook(historyInputWorkbook);
			projW = Workbook.getWorkbook(projectionInputWorkbook);
			
			// Get the sheets
			histSheet = histW.getSheet(0);
			projSheet = projW.getSheet(0);
						
			// Get global accountMap first
			getAccounts(histSheet, strACCOUNT_NAME_RANGE, strVOLUME_DATA_RANGE, strGROSS_MARGIN_RANGE, false);

			// Get warehouseMap, which is hashed by warehouses
			getWarehouseAccounts(histSheet);

			// Convert global warehouseMap into values represented by %
			// i.e., calculated via warehouseMap / accountMap
			convertWarehouseAccounts(warehouse_type.TO_PERCENTAGE_WAREHOUSE);

			// Get projected accounts
			getAccounts(projSheet, strPROJ_ACCOUNT_NAME_RANGE, strPROJ_VOLUME_DATA_RANGE, strPROJ_GROSS_MARGIN_RANGE, true);

			// Lastly get projected warehouseMap
			convertWarehouseAccounts(warehouse_type.TO_PROJECTION_WAREHOUSE);
			
			// Delete missing accounts from the projected warehouseMap
			deleteMissingAccountsFromMap();
			
			
			//Print warehouses & accounts
			//printAccounts(account_type.HISTORY_ACCOUNTS);
			//printAccounts(account_type.PROJECTION_ACCOUNTS);
			//printWarehouses();			
			
		} catch (BiffException e) {
			e.printStackTrace();
		}
	}

	public void deleteMissingAccountsFromMap() {
				
		// Let's delete missing entries from the gWarehouseMap 
		// This is the last step that needs to happen before projection is generated							
		Iterator<String[]> it = missingAccountList.iterator();
	    while (it.hasNext()) {
	    	
	    	String[] account = (String[])it.next();	    	
	    	String _warehouseName = account[0];
	    	String _accountName = account[1];
	    	
	    	Map<String, ArrayList<double[]>> aMap = gWarehouseMap.get(_warehouseName);
	    	aMap.remove(_accountName);
	    }		
	}
	
	public Sheet getHistSheet() {
		return histSheet;
	}

	public Sheet getProjSheet() {
		return projSheet;
	}

	public void loadConfig() {

		try {
			
			FileInputStream fstream = new FileInputStream(configDataPath);
			DataInputStream in = new DataInputStream(fstream);
			BufferedReader br = new BufferedReader(new InputStreamReader(in));
			String strLine = null;
			String regex = "=";
			
			while ((strLine = br.readLine()) != null)   {
				
				String[] theline = strLine.split(regex);
				
				String key = theline[0];
				String value = theline[1];
				
				if (configHash.containsKey(key)) {
					configHash.remove(key);
				}
				
				configHash.put(key, value);				
				System.out.println(key + " | " + value);
			}
			in.close();

			
		} catch (Exception e) {
			System.err.println("loadConfig Error: " + e.getMessage());
		}
	}				

	public double getWarehouseDataByMonth(warehouse_type type, double local, double global) {
		
		double retVal = 0f;
		
		if (type == warehouse_type.TO_PERCENTAGE_WAREHOUSE) 
			retVal = local/global;
		else
			retVal = local*global;
		
		return retVal; 
	}
	
	/*
	 * Take warehouseMap and divide by each account's total volume and gross margin per month
	 * to arrive at the map that now includes the % for each account per month 
	 */
	public void convertWarehouseAccounts(warehouse_type type) {
		
		Map<String, ArrayList<double[]>> globalAccountMap = gAccountMap;
		if (type == warehouse_type.TO_PROJECTION_WAREHOUSE) 
			globalAccountMap = gProjectedAccountMap;
			
		
		// Iterate through each warehouse
		for (Map.Entry<String, Map<String, ArrayList<double[]>>> warehouseEntry : gWarehouseMap.entrySet()) {			

			// Get warehouseName = 
			String warehouseName = warehouseEntry.getKey();
			
			// Get accountMap
			Map<String, ArrayList<double[]>> acctMap = warehouseEntry.getValue();

			// Iterate through each account belonging to each warehouse
			for (Map.Entry<String, ArrayList<double[]>> accountEntry : acctMap.entrySet()) {
				
				String accountName = accountEntry.getKey();				
				
				//System.out.println("For account: " + accountName);
				
				// GLOBAL: Get the array list from the global accountMap
				ArrayList<double[]> globalList = globalAccountMap.get(accountName);

				// LOCAL: Get the array list from local accountMap
				ArrayList<double[]> localList = accountEntry.getValue();				

				if (globalList == null) {
					
					System.out.println("Global list is NULL for " + warehouseName + " | " + accountName);
					
					// Add to the missing account list
					missingAccountList.add(new String[] {warehouseName, accountName});					
					
				} else {
					
					double[] globalVolumes = globalList.get(0);
					double[] localVolumes = localList.get(0);
					double[] globalMargins = globalList.get(1);
					double[] localMargins = localList.get(1);
					
					int len = localVolumes.length;
				
					if (globalVolumes.length != localVolumes.length || globalMargins.length != localMargins.length ) {						
						System.out.println("Global Size != Local Size");
					} else {
						
						//double[] localVolumesInPercent = new double[localVolumes.length];
						//double[] localMarginsInPercent = new double[globalMargins.length];
						
						// Convert local volume list & local margin list into %, and save
						// the results back to the original local list, to be put back to the 
						// original master map
						for (int i=0; i<len; i++) {

							if (globalVolumes[i] != 0) {
								localVolumes[i] = getWarehouseDataByMonth(
														type, 
														localVolumes[i], 
														globalVolumes[i]
													);
							} else { 
								localVolumes[i] = 0f;
							}
							//dataText += "\t" + getCustomFormat(localVolumes[i]);														
						}		
						
						for  (int j=0; j<len; j++) {
							
							if (globalMargins[j] != 0) {
								localMargins[j] = getWarehouseDataByMonth(
														type, 
														localMargins[j], 
														globalMargins[j]
													);		
							} else { 
								localMargins[j] = 0f;
							}					
							//dataText += "\t" + getCustomFormat(localMargins[j]);							
						}
												
						//dataText += 'n';

						// Put newly generated array lists into acctMap before next list iteration
						localList.set(0, localVolumes);
						localList.set(1, localMargins);
						
						acctMap.put(accountName, localList);
					} 
					
				} //end outer if 
				
				gWarehouseMap.put(warehouseName, acctMap);
			}					
		}
		
		//System.out.println(dataText);						
		//writeFile(dataText);
	}
	
	public String getCustomFormat(double value) {
		String pattern = "#0.0000";
		DecimalFormat myFormatter = new DecimalFormat(pattern);
		String output = myFormatter.format(value);		
		return output;
	}
	
	public void getWarehouseAccounts(Sheet sheet) {
		
		// This method loads all of the entries verbatim from the history.xls data
		// and generates a hash map of map of arraylist		
		Cell cell = sheet.getCell(configHash.get(strWAREHOUSE_DATA_STARTING_COORDINATE));
		int accountX = cell.getColumn();
		int accountY = cell.getRow();
		int numRows = sheet.getRows();
		String warehouseName = null;
		
		// Iterate through each warehouse name
		for (;accountY<numRows; accountY++) {
			
			Cell tmpCell = sheet.getCell(accountX, accountY);
			String tmpName = tmpCell.getContents();
			String accountName = sheet.getCell(accountX+1, accountY).getContents();
						
			if (tmpName.isEmpty() || tmpName == null || tmpName =="") {
								
				// check to see if its neighbor, i.e., accountName, has any stuff in it				
				if (accountName.isEmpty() == false) {
				
					// Let's pull the existing account map out and add this entry to it
					if (gWarehouseMap.containsKey(warehouseName)) {
						
						Map<String, ArrayList<double[]>> accountMap = gWarehouseMap.get(warehouseName);
												
						if (accountMap.isEmpty()) {
							System.out.println("ACCOUNT MAP IS NULL: " + warehouseName + " | " + accountName);
						}
						
						//===============
						// Get volumeData
						//===============
						String volumeRange = configHash.get(strVOLUME_DATA_RANGE);
						String[] vRange = volumeRange.split(":");
						int startingVolumeX = sheet.getCell(vRange[0]).getColumn();		
						int endingVolumeX = sheet.getCell(vRange[1]).getColumn();
						double[] volumeData = getVolumeDataPerAccount(sheet, startingVolumeX, endingVolumeX, accountY);
												
						//===============				
						// Get gross margins
						//===============
						String marginRange = configHash.get(strGROSS_MARGIN_RANGE);
						String[] mRange = marginRange.split(":");
						int startingMarginX = sheet.getCell(mRange[0]).getColumn();		
						int endingMarginX = sheet.getCell(mRange[1]).getColumn();
						double[] marginData = getMarginDataPerAccount(sheet, startingMarginX, endingMarginX, accountY);

						// Let's update the list
						ArrayList<double[]> list = accountMap.get(accountName);

						if (list == null) {							
							//System.out.println("here: " + Arrays.toString(volumeData));
							list = new ArrayList<double[]>();
							list.add(0, volumeData);
							list.add(1, marginData);
						} else {
							list.set(0, volumeData);
							list.set(1, marginData);
						}
																		
						// Add updated list to accountMap
						accountMap.put(accountName, list);												

						// put this into map
						gWarehouseMap.put(warehouseName, accountMap);

						//System.out.println("HASHING: " + warehouseName + " with " + accountName);
						
					} else {
						System.out.println("WAREHOUSE NOT IN EXISTENCE: " + warehouseName);
					}
					
	
				}				
				
			} else if (tmpName.matches("(?i).*Total.*")) {
			
				//warehouseName = null;				
				//System.out.println("TOTAL: " + tmpName);
				
			} else {

				warehouseName = tmpName;
								
				// check to see if its neighbor, i.e., accountName, has any stuff in it
				// If so, let's create an account map here
				if (accountName.isEmpty() == false) {
				
					Map<String, ArrayList<double[]>> accountMap = new HashMap<String, ArrayList<double[]>>(); 
					ArrayList<double[]> list = new ArrayList<double[]>();
					
					// Add to list
					//===============
					// Get volumeData
					//===============
					String volumeRange = configHash.get(strVOLUME_DATA_RANGE);
					String[] vRange = volumeRange.split(":");
					int startingVolumeX = sheet.getCell(vRange[0]).getColumn();		
					int endingVolumeX = sheet.getCell(vRange[1]).getColumn();
					double[] volumeData = getVolumeDataPerAccount(sheet, startingVolumeX, endingVolumeX, accountY);
					list.add(0, volumeData);
											
					//===============				
					// Get gross margins
					//===============
					String marginRange = configHash.get(strGROSS_MARGIN_RANGE);
					String[] mRange = marginRange.split(":");
					int startingMarginX = sheet.getCell(mRange[0]).getColumn();		
					int endingMarginX = sheet.getCell(mRange[1]).getColumn();
					double[] marginData = getMarginDataPerAccount(sheet, startingMarginX, endingMarginX, accountY);
					list.add(1, marginData);
					
					// Add updated list to accountMap
					accountMap.put(accountName, list);
					
					// put this into warehouse map
					gWarehouseMap.put(warehouseName, accountMap);
					
					//System.out.println("CREATING " + warehouseName + " with " + accountName);
					
				} else {
					System.out.println("ACCOUNT NAME IS EMPTY for WAREHOUSE: " + warehouseName);
				}
				
			}
		}		
	}
	
	
	public void getAccounts(Sheet sheet, String accountRange, String vRange, String mRange, boolean isProjection) {
		
		// Determine the starting X, Y for all account names
		String strAccountRange = configHash.get(accountRange);
		String[] acctRangeArray = strAccountRange.split(":");
		Cell cell = sheet.getCell(acctRangeArray[0]);
				
		int accountX = cell.getColumn();
		int accountY = cell.getRow();	
		int numRows = sheet.getCell(acctRangeArray[1]).getRow();	
				
		// Iterate through each account name
		for (;accountY<=numRows; accountY++) {

			Cell tmpCell = sheet.getCell(accountX, accountY);
			String cellContents = tmpCell.getContents();
			
			if (cellContents.isEmpty() || cellContents == null || cellContents =="") {
				// Don't add to hash
			} else {				

				String accountName = sheet.getCell(accountX, accountY).getContents();
				
				//System.out.println("Account = " + accountName);

				// Let's get the volumes and gross margins for this account
				getAccountData(accountName, sheet, accountY, vRange, mRange, isProjection);
				
			}
		}
		
				
	}
	
	
	public void getAccountData(String name, Sheet sheet, int rowY, String vRange, String mRange, boolean isProjection) {

		Map<String, ArrayList<double[]>> map;		
		if (isProjection) 
			map = gProjectedAccountMap;
		else 
			map = gAccountMap;
		
		//===============
		// Get volumeData
		//===============a
		String volumeRange = configHash.get(vRange);
		String[] vRangeArray = volumeRange.split(":");
		int startingVolumeX = sheet.getCell(vRangeArray[0]).getColumn();		
		int endingVolumeX = sheet.getCell(vRangeArray[1]).getColumn();
		double[] volumeData = getVolumeDataPerAccount(sheet, startingVolumeX, endingVolumeX, rowY);

		//===============				
		// Get gross margins
		//===============
		String marginRange = configHash.get(mRange);
		String[] mRangeArray = marginRange.split(":");
		int startingMarginX = sheet.getCell(mRangeArray[0]).getColumn();		
		int endingMarginX = sheet.getCell(mRangeArray[1]).getColumn();
		double[] marginData = getMarginDataPerAccount(sheet, startingMarginX, endingMarginX, rowY);
				
		// Check to see if account name is already in the map
		if (map.containsKey(name)) {
			
			ArrayList<double[]> list = map.get(name);
			
			if (list == null) {
				
				list = new ArrayList<double[]>();
				list.add(0, volumeData);
				list.add(1, marginData);
				
			} else {	
				
				double[] oldVolumeData = (double[])list.get(0);
				double[] newVolumeData = addData(oldVolumeData, volumeData);				
				double[] oldMarginData = (double[])list.get(1);
				double[] newMarginData = addData(oldMarginData, marginData);
				list.set(0, newVolumeData);
				list.set(1, newMarginData);				
			}
			
			map.put(name, list);			
			
		} else {

			System.out.println("Account Map w/ new AccountName: " + name);
			
			// Create an array list 
			ArrayList<double[]> list = new ArrayList<double[]>();
			list.add(0, volumeData);
			list.add(1, marginData);
			
			map.put(name, list);	
		}
	}
	
	public void printWarehouses() {
		
		String dataText = "";
		
		// Print warehouses
		for (Map.Entry<String, Map<String, ArrayList<double[]>>> masterEntry : gWarehouseMap.entrySet()) {			
			
			dataText += masterEntry.getKey() + " :\n";
			
			// Print accounts
			Map<String, ArrayList<double[]>> accountMap = masterEntry.getValue();

			for (Map.Entry<String, ArrayList<double[]>> accountEntry : accountMap.entrySet()) {
				
				String accountName = accountEntry.getKey();				
				dataText += "\t" + accountName;	
				
				ArrayList<double[]> list = accountEntry.getValue();
								
				// print volumes
				double[] volumes = list.get(0);
				for (int i=0; i<volumes.length; i++) {
					dataText += "\t" + volumes[i];				
				}
				
				// print gross margins
				double[] margins = list.get(1);
				for (int i=0; i<margins.length; i++) {
					dataText += "\t" + margins[i];				
				}
				dataText += "\n";
			}					
		}
				
		System.out.println(dataText);
		
		writeFile(warehousesPath, dataText);
		
	}
	
	public void printAccounts(account_type accountType) {
		
		String dataText = "";

		String path = projAccountPath;
		Map<String, ArrayList<double[]>> map = gProjectedAccountMap;
				
		if (accountType==account_type.HISTORY_ACCOUNTS) {
			path = histAccountPath;
			map = gAccountMap;
		}		
		
		// Print volumes
		for (Map.Entry<String, ArrayList<double[]>> masterEntry : map.entrySet()) {			
			
			dataText += masterEntry.getKey() + " :";
			
			// Print volumes
			ArrayList<double[]> list = masterEntry.getValue();
			double[] volumes = (double[])list.get(0);		
			for (int i=0; i<volumes.length; i++) {
				dataText += "\t" + volumes[i];				
			}

			double[] margins = (double[])list.get(1);
			for (int i=0; i<margins.length; i++) {
				dataText += "\t" + margins[i];				
			}
						
			dataText += "\n";
		}
				
		System.out.println(dataText);
		System.out.println("TOTAL ACCOUNTS: " + map.size());	
				
		writeFile(path, dataText);
	}
	
	public double[] addData(double[] oldArray, double[] newArray) {
			
		for (int i = 0; i<oldArray.length; i++) {
			oldArray[i] += newArray[i];
		}
		
		return oldArray;
	}
	
	public double[] getMarginDataPerAccount(Sheet sheet, int start, int end, int y) {
		
		int delta = end - start + 1;
		
		double[] marginArray = new double[delta];
		int i = 0;
		for (int index=start; index<=end; index++) {
			
			String marginData = sanitizeString(sheet.getCell(index, y).getContents());
			if (marginData.isEmpty() == false) {
				marginArray[i] = Double.parseDouble(marginData);
			} else {
				marginArray[i] = 0;
			}
			
			//System.out.println("Margin Data: " + marginArray[i]);
			
			i++;			
		}
		
		return marginArray;
	}
	
	public double[] getVolumeDataPerAccount(Sheet sheet, int start, int end, int y) {
		
		int delta = end - start + 1;

		double[] volumeArray = new double[delta];
		int i = 0;
		for (int index=start; index<=end; index++) {
			
			String volumeData = sanitizeString(sheet.getCell(index, y).getContents());
			
			if (volumeData.isEmpty() == false) {
								
				volumeArray[i] = Double.parseDouble(volumeData);
			} else {
				volumeArray[i] = 0;
			}
			
			//System.out.println("Volume Data: " + volumeArray[i]);
			
			i++;			
		}
		
		return volumeArray;
	}
 	
	public String sanitizeString(String str) {
		
		String retVal = str;
		
		// get rid of all commas, periods, $, etc		
		retVal = retVal.replace("\"$\"", "");
		
		return retVal.replace(",", "");
	}

	
	public void writeFile(String path, String text) {
				
		System.out.println("text = " + text);
		
		try {
			
			// Delete file is already exist
			File file = new File(path);			
			if (file.exists()) {
				file.delete();
			} 
			
			// Create file
			FileWriter fstream = new FileWriter(path);
			BufferedWriter out = new BufferedWriter(fstream);
			out.write(text);
			
			//Close the output stream
			out.close();
			
		} catch (Exception e){//Catch exception if any
			System.err.println("Error: " + e.getMessage());
		}		
	}
	
	public String[] getLabels(Sheet sheet, String range) {
		
		String[] retVal;
		String strRange = configHash.get(range);
		String[] rangeArray = strRange.split(":");
		
		Cell cellStart = sheet.getCell(rangeArray[0]);
		Cell cellEnd = sheet.getCell(rangeArray[1]);
		
		int y = cellStart.getRow();
		int startX = cellStart.getColumn();
		int endX = cellEnd.getColumn();
		
		ArrayList<String> list = new ArrayList<String>();
		retVal = new String[endX-startX+1];
		for (int i=startX; i<=endX; i++) {			
			list.add(sheet.getCell(i,y).getContents());						
		}
		
		retVal = (String[])list.toArray(new String[list.size()]);

		//System.out.println(Arrays.toString(retVal));
		
		return retVal;		
	}
	
	
	public static void main(String[] args) throws WriteException, IOException {
		
		long start = System.currentTimeMillis();
		
		BudgetAnalysis analysis = new BudgetAnalysis();

		// First let's load config
		analysis.loadConfig();
		analysis.setInputFiles(configHash.get("HIST_DATA_FILE_NAME"), configHash.get("PROJ_DATA_FILE_NAME"));
		analysis.read();
		
		WriteExcel write = new WriteExcel(analysis);
		write.setOutputFiles(outputDataPath, missingDataPath);
		write.write();
		write.writeMissingData();
		
		System.out.println("This reports takes " + (System.currentTimeMillis()-start)/1000F + " sec to generate!");
		
	}

	
}
