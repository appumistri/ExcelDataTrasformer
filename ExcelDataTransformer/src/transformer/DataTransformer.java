package transformer;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Paths;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.widgets.Combo;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Event;
import org.eclipse.swt.widgets.FileDialog;
import org.eclipse.swt.widgets.Group;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.widgets.Listener;
import org.eclipse.swt.widgets.MessageBox;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.swt.widgets.Text;
import org.eclipse.wb.swt.SWTResourceManager;

public class DataTransformer {

	private Shell shell;

	/**
	 * Launch the application.
	 * 
	 * @param args
	 */
	public static void main(String[] args) {
		try {
			DataTransformer window = new DataTransformer();
			window.open();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * Open the window.
	 */
	public void open() {
		Display display = Display.getDefault();
		createContents();
		shell.open();
		shell.layout();
		while (!shell.isDisposed()) {
			if (!display.readAndDispatch()) {
				display.sleep();
			}
		}
	}

	/**
	 * Create contents of the window.
	 */
	protected void createContents() {
		shell = new Shell();
		shell.setImage(SWTResourceManager.getImage(DataTransformer.class, "/transformer/resources/CES_logo.png"));
		shell.setSize(635, 273);
		shell.setText("Excel Data Transformer");

		Group inputGroup = new Group(shell, SWT.NONE);
		inputGroup.setText("Inputs");
		inputGroup.setBounds(10, 10, 599, 214);

		Label leadsExcelLbl = new Label(inputGroup, SWT.NONE);
		leadsExcelLbl.setText("Leads Excel");
		leadsExcelLbl.setBounds(10, 25, 150, 30);

		Text leadsExcelPathTxt = new Text(inputGroup, SWT.BORDER);
		leadsExcelPathTxt.setToolTipText("Excel Path");
		leadsExcelPathTxt.setBounds(200, 25, 304, 25);

		Button browseLeadsExcel = new Button(inputGroup, SWT.NONE);
		browseLeadsExcel.setBounds(507, 25, 82, 25);
		browseLeadsExcel.setText("Browse");
		browseLeadsExcel.addListener(SWT.Selection, new Listener() {

			@Override
			public void handleEvent(Event arg0) {
				String filePath = selectFile();
				if (filePath != null) {
					leadsExcelPathTxt.setText(filePath);
				}
			}
		});

		Label SheetNameLbl = new Label(inputGroup, SWT.NONE);
		SheetNameLbl.setText("Sheet Name");
		SheetNameLbl.setBounds(10, 62, 181, 27);

		Text sheetNameTxt = new Text(inputGroup, SWT.BORDER);
		sheetNameTxt.setToolTipText("Sheet Name");
		sheetNameTxt.setBounds(200, 60, 304, 25);

		String[] items = new String[] { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", };
		Combo columnNumberCmb = new Combo(inputGroup, SWT.READ_ONLY);
		columnNumberCmb.setToolTipText("Comumn Number");
		columnNumberCmb.setBounds(200, 94, 181, 23);
		columnNumberCmb.setItems(items);
		columnNumberCmb.select(0);

		Button transformBtn = new Button(inputGroup, SWT.NONE);
		transformBtn.setBounds(200, 161, 181, 33);
		transformBtn.setText("Transform");

		Label excelColumn = new Label(inputGroup, SWT.NONE);
		excelColumn.setText("Column Number");
		excelColumn.setBounds(10, 93, 181, 27);

		Label delimeterLbl = new Label(inputGroup, SWT.NONE);
		delimeterLbl.setBounds(10, 126, 55, 15);
		delimeterLbl.setText("Delimeter");

		Text delimeterTxt = new Text(inputGroup, SWT.BORDER);
		delimeterTxt.setBounds(200, 123, 181, 21);
		transformBtn.addListener(SWT.Selection, new Listener() {

			@Override
			public void handleEvent(Event arg0) {
				String excelPath = leadsExcelPathTxt.getText();
				String sheetName = sheetNameTxt.getText();
				int columnNumber = columnNumberCmb.getSelectionIndex();
				String delimeter = delimeterTxt.getText().trim();

				if (excelPath.isEmpty() || sheetName.isEmpty()) {
					emptyInputWarning();
				} else {
					transform(excelPath, sheetName, columnNumber, delimeter);
				}
			}
		});
	}

	public String selectFile() {
		FileDialog dialog = new FileDialog(shell, SWT.OPEN);
		dialog.setFilterExtensions(new String[] { "*.xlsx", "*.xls" });
		dialog.setFilterPath(System.getProperty("uder.dir"));
		String filePath = dialog.open();
		return filePath;
	}

	public void emptyInputWarning() {
		MessageBox messageBox = new MessageBox(shell, SWT.ICON_WARNING | SWT.OK);
		messageBox.setText("Warning");
		messageBox.setMessage("Please provide necessary input in fields.");
		messageBox.open();
	}

	public void transform(String excelPath, String sheetName, int columnNumber, String delimeter) {
		try {
			Workbook inWorkbook = new XSSFWorkbook(excelPath);
			Sheet inSheet = inWorkbook.getSheet(sheetName);
			int lastRow = inSheet.getLastRowNum();

			Workbook outWorkbook = new XSSFWorkbook();
			Sheet outSheet = outWorkbook.createSheet();

			int outColumnNumber = 0, outRowNumber = 0;

			for (int i = 0; i <= lastRow; i++) {
				Row inRow = inSheet.getRow(i);
				Cell inCell = inRow.getCell(columnNumber);
				String text = getCellValue(inCell);

				if (text == null) {
					break;
				}

				Row outRow = outSheet.getRow(outRowNumber);
				if (outRow == null) {
					outRow = outSheet.createRow(outRowNumber);
				}

				Cell outCell = outRow.createCell(outColumnNumber);
				outCell.setCellValue(text);
				outColumnNumber++;

				System.out.println("Row:" + outRowNumber);
				System.out.println("Column:" + outColumnNumber);

				if (delimeter.equals(text)) {
					outColumnNumber = 0;
					outRowNumber++;
				}
			}

			File inFile = new File(excelPath);
			String outFileName = Paths.get(System.getProperty("user.dir"), "Transformed_" + inFile.getName())
					.toString();
			FileOutputStream fileOutStrem = new FileOutputStream(outFileName);
			outWorkbook.write(fileOutStrem);
			fileOutStrem.close();
			outWorkbook.close();
			inWorkbook.close();

		} catch (IOException e) {
			throw new RuntimeException(e);
		}
	}

	public String getCellValue(Cell cell) {
		String cellvalue = null;
		try {
			switch (cell.getCellType()) {
			case STRING:
				cellvalue = cell.getStringCellValue().trim();
				break;
			case NUMERIC:
				cellvalue = String.valueOf(cell.getNumericCellValue());
				break;
			case _NONE:
				cellvalue = cell.getStringCellValue().trim();
				break;
			case BLANK:
				cellvalue = "";
				break;
			case BOOLEAN:
				cellvalue = String.valueOf(cell.getBooleanCellValue());
				break;
			case ERROR:
				cellvalue = String.valueOf(cell.getErrorCellValue());
				break;
			case FORMULA:
				cellvalue = String.valueOf(cell.getCellFormula());
				break;
			default:
				throw new RuntimeException("Unknown Cell type: " + cell.getCellType());
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return cellvalue;
	}
}
