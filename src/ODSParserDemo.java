import java.io.File;
import java.io.IOException;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import org.jopendocument.dom.spreadsheet.*;



public class ODSParserDemo {
	
	public static void main(String[] args) {
		
		NumberFormat money = NumberFormat.getCurrencyInstance();
		
		//ODF filename here
		String filename = "test.ods";
		
		//Similar to the HSSF demo, we can read in everything into lists of
		//lists, but it'd be better to abstract everything from here while
		//the file is being read so exceptions can be handled.
		
		List columns = new ArrayList();
		
		try {
			File f = new File(filename);
			
			//We'll assume everything we want is on the first sheet:
			final Sheet sheet = SpreadSheet.createFromFile(f).getSheet(0);
			
			System.out.println("Sheet contains " + sheet.getColumnCount() + " columns.");
			System.out.println("Sheet contains " + sheet.getRowCount() + " rows.");
			
			//Check that the spreadsheet has the required columns:
			if (sheet.getColumnCount() < 7) {
				System.err.println("Spreadsheet does not have expected number of columns! ");
				//Throw custom exception here
			}
						
			//I'd like my first row to be a title, so to match the column index
			//to the title we'll make an hashmap:
			Map<String, Integer> columnDictionary = new TreeMap<String, Integer>();
			for (int col = 0; col < sheet.getColumnCount(); col++) {
				//map the value of each column in row 0 to corresponding column index
				String cellText = sheet.getImmutableCellAt(col, 0).getTextValue();
				columnDictionary.put(cellText.toLowerCase(), col);
			}
			
			//Check that the required column names are set:
			String[] requiredColumns = {"food", "category", "price", "quantity", 
					"description", "size", "special order"};
			for (String key : requiredColumns) {
				if (!columnDictionary.containsKey(key)) {
					System.err.println("Did not find required column name \""+key+"\"");
					//Throw custom exception here
				}
			}
			
			//Parse the arguments out here
			List<String> foodList = ColumnToList(sheet, columnDictionary.get("food")); 
			List<String> categoryList = ColumnToList(sheet, columnDictionary.get("category"));
			List<String> priceList = ColumnToList(sheet, columnDictionary.get("price"));
			List<String> quantityList = ColumnToList(sheet, columnDictionary.get("quantity"));
			List<String> descriptionList = ColumnToList(sheet, columnDictionary.get("description"));
			//...
			
			/*Alternatively, you can read one row at a time and build each food item that way:
			 *   List<String> row = RowToList(sheet, 1);
			 *   FoodItem = new FoodItem(row[columnDictionary.get("food"), ...);
			 *   ...
			 */
			
			for (int row = 1; row < sheet.getRowCount(); row++) {
				System.out.println(foodList.get(row-1) + "\t(" + 
						money.format(Float.parseFloat(priceList.get(row-1))) + ") × " + 
						quantityList.get(row-1));
				System.out.println("\tCategory: " + categoryList.get(row-1));
				System.out.println("\t" + descriptionList.get(row-1));
				System.out.println("\n");
			}
			
			
		} catch (IOException e) {
			e.printStackTrace();
		}
		
	}
	
	
	/**
	 * Gets an entire column from a sheet as one list.
	 * @param sheet
	 * @param column
	 * @return List<String>
	 */
	public static List<String> ColumnToList(Sheet sheet, int column) {
		List<String> ret = new ArrayList<String>();
		for (int row = 1; row < sheet.getRowCount(); row++) {
			String cellText = sheet.getImmutableCellAt(column, row).getTextValue();
			ret.add(cellText);
		}
		return ret;
	}
	
	/**
	 * Gets an entire row from one sheet as a list.
	 * @param sheet
	 * @param row
	 * @return List<String>
	 */
	public static List<String> RowToList(Sheet sheet, int row) {
		List<String> ret = new ArrayList<String>();
		for (int col = 0; col < sheet.getColumnCount(); col++)
			ret.add(sheet.getImmutableCellAt(col, row).getTextValue());
		return ret;
	}

}
