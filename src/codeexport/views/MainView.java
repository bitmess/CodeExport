package codeexport.views;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Collection;
import java.util.List;

import javax.inject.Inject;

import org.apache.commons.io.FileUtils;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.eclipse.jface.dialogs.MessageDialog;
import org.eclipse.jface.viewers.ComboViewer;
import org.eclipse.swt.SWT;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.layout.FillLayout;
import org.eclipse.swt.layout.GridData;
import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.widgets.Combo;
import org.eclipse.swt.widgets.Composite;
import org.eclipse.swt.widgets.DirectoryDialog;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Event;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.widgets.Text;
import org.eclipse.ui.IActionBars;
import org.eclipse.ui.IWorkbench;
import org.eclipse.ui.part.ViewPart;

/**
 * This sample class demonstrates how to plug-in a new workbench view. The view
 * shows data obtained from the model. The sample creates a dummy model on the
 * fly, but a real implementation would connect to the model available either in
 * this or another plug-in (e.g. the workspace). The view is connected to the
 * model using a content provider.
 * <p>
 * The view uses a label provider to define how model objects should be
 * presented in the view. Each view can present the same model objects using
 * different labels and icons, if needed. Alternatively, a single label provider
 * can be shared between views in order to ensure that objects of the same type
 * are presented in the same way everywhere.
 * <p>
 */

public class MainView extends ViewPart {

	public MainView() {
	}

	/**
	 * The ID of the view as specified by the extension.
	 */
	public static final String ID = "codeexport.views.MainView";

	@Inject
	IWorkbench workbench;
	private Text inputDirText;

	private Button inputButton;

	private XWPFDocument document;
	private Text exportFilePathText;

	private Button exportButton;

	private Button exportFileButton;
	private Label label_1;
	private Label label_2;
	private Combo combo;
	private ComboViewer comboViewer;
	private Button addFileTypeButton;
	private Composite composite;
	private Button resetFileTypeButton;

	/**
	 * 2007word文档创建
	 * 
	 * @throws IOException
	 */
	private void createWord2007() throws IOException {
		

		IActionBars bars = getViewSite().getActionBars();
		bars.getStatusLineManager().setMessage("处理中");

		String[] fileFilter = combo.getItems();

		Display.getDefault().asyncExec(() -> {

			Collection<File> pathes = FileUtils.listFiles(new File(inputDirText.getText()), fileFilter, true);
			String exportPath = exportFilePathText.getText();
			exportPath = exportPath + File.separator + System.currentTimeMillis() + ".docx";

//		txt
//		File f = new File(exportPath);
//		for (File file : pathes) {
//			StringBuffer sb = new StringBuffer();
//			List<String> readLines = FileUtils.readLines(file, "UTF-8");
//			for (int i = 0; i < readLines.size(); i++) {
//				String string = readLines.get(i);
//				sb.append(string).append(System.getProperty("line.separator"));
//			}
//			
//			String result = sb.toString();
//			//去掉注释
//			///\*[\w\W]*?\*/
//			////.*
//			result = result.replaceAll("\\*[\\w\\W]*?\\*/", "");
//			result = result.replaceAll("//.*", "");
//			
//			//去掉空白行(?m)^\\s*$(\\n|\\r\\n)
//			result = result.replaceAll("(?m)^\\s*$(\\n|\\r\\n)", "");
//			
//			FileUtils.write(f, result, "UTF-8", true);
//		}

//		word
			// Write the Document in file system
//		exportPath
			FileOutputStream out;
			try {
				out = new FileOutputStream(new File(exportPath));
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
				Display.getDefault().syncExec(() -> {

					bars.getStatusLineManager().setMessage("错误");
					updateUIByEnable(true);

				});
				return;
			}
//		FileOutputStream out = new FileOutputStream(new File("/Users/jv/Desktop/create_toc.docx"));

			document = new XWPFDocument();// 文档

			// 段落
			for (File file : pathes) {
				XWPFParagraph firstParagraph = document.createParagraph();
				firstParagraph.setAlignment(ParagraphAlignment.LEFT);

				XWPFRun run = firstParagraph.createRun();
				StringBuffer sb = new StringBuffer();
				List<String> readLines;
				try {
					readLines = FileUtils.readLines(file, "UTF-8");
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
					Display.getDefault().syncExec(() -> {
						bars.getStatusLineManager().setMessage("错误");
						updateUIByEnable(true);
					});
					return;
				}
				for (int i = 0; i < readLines.size(); i++) {
					String string = readLines.get(i);
					sb.append(string).append(System.getProperty("line.separator"));
				}

				String result = sb.toString();
				// 去掉注释
				/// \*[\w\W]*?\*/
				//// .*
				result = result.replaceAll("\\*[\\w\\W]*?\\*/", "");
				result = result.replaceAll("//.*", "");

				// 去掉空白行(?m)^\\s*$(\\n|\\r\\n)
				result = result.replaceAll("(?m)^\\s*$(\\n|\\r\\n)", "");

				if (result.contains("\n")) {
					String[] lines = result.split("\n");
					run.setText(lines[0], 0); // set first line into XWPFRun
					for (int i = 1; i < lines.length; i++) {
						// add break and insert new text
						run.addBreak();
						run.setText(lines[i]);
					}
				} else {
					run.setText(result, 0);
				}
			}

			document.createTOC();

			try {
				document.write(out);
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			try {
				out.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}

			Display.getDefault().syncExec(() -> {
				bars.getStatusLineManager().setMessage("成功");
				updateUIByEnable(true);
			});

		});

	}

	private void updateUIByEnable(boolean enabled) {
		addFileTypeButton.setEnabled(enabled);
		resetFileTypeButton.setEnabled(enabled);
		combo.setEnabled(enabled);
		inputDirText.setEnabled(enabled);
		exportFilePathText.setEnabled(enabled);
		exportButton.setEnabled(enabled);
		inputButton.setEnabled(enabled);
		exportFileButton.setEnabled(enabled);
	}

	private void startExport() throws IOException {

		if (inputDirText.getText().isEmpty()) {

			MessageDialog.openError(null, "提示", "根目录为空");

			inputButton.notifyListeners(SWT.Selection, new Event());

			return;
		}

		if (exportFilePathText.getText().isEmpty()) {

			MessageDialog.openError(null, "提示", "导出目录为空");
			exportButton.notifyListeners(SWT.Selection, new Event());

			return;
		}

		if (combo.getItems().length == 0) {
			MessageDialog.openError(null, "提示", "文件类型不能为空");
			combo.setFocus();
			return;
		}

		updateUIByEnable(false);

		createWord2007();
	}

	@Override
	public void createPartControl(Composite parent) {
		parent.setLayout(new GridLayout(3, false));

		Label fileLabel = new Label(parent, SWT.NONE);
		fileLabel.setLayoutData(new GridData(SWT.RIGHT, SWT.CENTER, false, false, 1, 1));
		fileLabel.setText("目录路径");

		inputDirText = new Text(parent, SWT.BORDER);
		inputDirText.setLayoutData(new GridData(SWT.FILL, SWT.CENTER, true, false, 1, 1));

		inputButton = new Button(parent, SWT.NONE);
		inputButton.setLayoutData(new GridData(SWT.CENTER, SWT.CENTER, false, false, 1, 1));
		inputButton.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				DirectoryDialog dd = new DirectoryDialog(workbench.getDisplay().getActiveShell());
				String path = dd.open();
				if (path != null && !path.isEmpty()) {
					inputDirText.setText(path);
				}
			}
		});
		inputButton.setText("项目根目录");

		Label label = new Label(parent, SWT.NONE);
		label.setLayoutData(new GridData(SWT.RIGHT, SWT.CENTER, false, false, 1, 1));
		label.setText("导出文件目录");

		exportFilePathText = new Text(parent, SWT.BORDER);
		exportFilePathText.setLayoutData(new GridData(SWT.FILL, SWT.CENTER, true, false, 1, 1));

		exportButton = new Button(parent, SWT.NONE);
		exportButton.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				DirectoryDialog dd = new DirectoryDialog(workbench.getDisplay().getActiveShell());
				String path = dd.open();
				if (path != null && !path.isEmpty()) {
					exportFilePathText.setText(path);
				}
			}
		});
		exportButton.setLayoutData(new GridData(SWT.CENTER, SWT.CENTER, false, false, 1, 1));
		exportButton.setText("导出目录");

		label_2 = new Label(parent, SWT.NONE);
		label_2.setLayoutData(new GridData(SWT.RIGHT, SWT.CENTER, false, false, 1, 1));
		label_2.setText("文件类型");

		comboViewer = new ComboViewer(parent, SWT.NONE);
		combo = comboViewer.getCombo();
		combo.setLayoutData(new GridData(SWT.FILL, SWT.CENTER, true, false, 1, 1));

		composite = new Composite(parent, SWT.NONE);
		composite.setLayout(new FillLayout(SWT.VERTICAL));
		composite.setLayoutData(new GridData(SWT.FILL, SWT.FILL, false, false, 1, 1));

		addFileTypeButton = new Button(composite, SWT.NONE);
		addFileTypeButton.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				String s = combo.getText();
				if (s != null && !s.isEmpty()) {
					combo.add(s);
					combo.setText("");
				}
			}
		});
		addFileTypeButton.setText("添加文件类型");

		resetFileTypeButton = new Button(composite, SWT.NONE);
		resetFileTypeButton.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				combo.setText("");
				combo.removeAll();
			}
		});
		resetFileTypeButton.setText("重置文件类型");

		label_1 = new Label(parent, SWT.NONE);
		label_1.setLayoutData(new GridData(SWT.CENTER, SWT.CENTER, false, false, 1, 1));

		exportFileButton = new Button(parent, SWT.NONE);
		exportFileButton.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				try {
					startExport();
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			}
		});
		exportFileButton.setLayoutData(new GridData(SWT.CENTER, SWT.CENTER, false, false, 1, 1));
		exportFileButton.setText("开始");
		new Label(parent, SWT.NONE);

	}

	@Override
	public void setFocus() {
		inputButton.setFocus();
	}
}
