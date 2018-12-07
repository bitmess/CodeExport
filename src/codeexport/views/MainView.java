package codeexport.views;


import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.inject.Inject;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.eclipse.jface.action.Action;
import org.eclipse.jface.action.IMenuListener;
import org.eclipse.jface.action.IMenuManager;
import org.eclipse.jface.action.IToolBarManager;
import org.eclipse.jface.action.MenuManager;
import org.eclipse.jface.action.Separator;
import org.eclipse.jface.dialogs.MessageDialog;
import org.eclipse.jface.viewers.ArrayContentProvider;
import org.eclipse.jface.viewers.DoubleClickEvent;
import org.eclipse.jface.viewers.IDoubleClickListener;
import org.eclipse.jface.viewers.IStructuredSelection;
import org.eclipse.jface.viewers.ITableLabelProvider;
import org.eclipse.jface.viewers.LabelProvider;
import org.eclipse.jface.viewers.TableViewer;
import org.eclipse.swt.SWT;
import org.eclipse.swt.graphics.Image;
import org.eclipse.swt.widgets.Composite;
import org.eclipse.swt.widgets.DirectoryDialog;
import org.eclipse.swt.widgets.FileDialog;
import org.eclipse.swt.widgets.Menu;
import org.eclipse.ui.IActionBars;
import org.eclipse.ui.ISharedImages;
import org.eclipse.ui.IWorkbench;
import org.eclipse.ui.IWorkbenchActionConstants;
import org.eclipse.ui.PlatformUI;
import org.eclipse.ui.part.ViewPart;
import org.eclipse.swt.layout.FillLayout;
import org.eclipse.swt.layout.FormLayout;
import org.eclipse.swt.widgets.Group;
import org.eclipse.swt.layout.FormData;
import org.eclipse.swt.layout.FormAttachment;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.custom.CCombo;
import org.eclipse.swt.widgets.Combo;
import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.widgets.Text;
import org.eclipse.swt.layout.GridData;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;

/**
 * This sample class demonstrates how to plug-in a new
 * workbench view. The view shows data obtained from the
 * model. The sample creates a dummy model on the fly,
 * but a real implementation would connect to the model
 * available either in this or another plug-in (e.g. the workspace).
 * The view is connected to the model using a content provider.
 * <p>
 * The view uses a label provider to define how model
 * objects should be presented in the view. Each
 * view can present the same model objects using
 * different labels and icons, if needed. Alternatively,
 * a single label provider can be shared between views
 * in order to ensure that objects of the same type are
 * presented in the same way everywhere.
 * <p>
 */

public class MainView extends ViewPart {
	public MainView() {
	}

	/**
	 * The ID of the view as specified by the extension.
	 */
	public static final String ID = "codeexport.views.MainView";

	@Inject IWorkbench workbench;
	private Text text;
	 

	/**
	 * 2007word文档创建
	 * @throws IOException 
	 */
	public void createWord2007() throws IOException {
		XWPFDocument document= new XWPFDocument();

        //Write the Document in file system
        FileOutputStream out = new FileOutputStream(new File("/Users/jv/Desktop/create_toc.docx"));

        //添加标题
        XWPFParagraph titleParagraph = document.createParagraph();

        //设置段落居中
        titleParagraph.setAlignment(ParagraphAlignment.CENTER);

        XWPFRun titleParagraphRun = titleParagraph.createRun();
        titleParagraphRun.setText("Java PoI");
        titleParagraphRun.setColor("000000");
        titleParagraphRun.setFontSize(20);

        //段落
        XWPFParagraph firstParagraph = document.createParagraph();
        firstParagraph.setStyle("Heading1");
        XWPFRun run = firstParagraph.createRun();
        run.setText("段落1。");
        run.setColor("696969");
        run.setFontSize(18);


        //段落
        XWPFParagraph firstParagraph1 = document.createParagraph();
        firstParagraph.setStyle("Heading1");
        XWPFRun run1 = firstParagraph1.createRun();
        run1.setText("段落2");
        run1.setColor("696969");
        run1.setFontSize(16);

        document.createTOC();

        document.write(out);
        out.close();
	}


	@Override
	public void createPartControl(Composite parent) {
		parent.setLayout(new GridLayout(3, false));
		
		Label fileLabel = new Label(parent, SWT.NONE);
		fileLabel.setLayoutData(new GridData(SWT.RIGHT, SWT.CENTER, false, false, 1, 1));
		fileLabel.setText("目录路径");
		
		text = new Text(parent, SWT.BORDER);
		text.setLayoutData(new GridData(SWT.FILL, SWT.CENTER, true, false, 1, 1));
		
		Button btnNewButton = new Button(parent, SWT.NONE);
		btnNewButton.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				DirectoryDialog dd = new DirectoryDialog(workbench.getDisplay().getActiveShell());
				String path = dd.open();
				text.setText(path);
			}
		});
		btnNewButton.setText("选择目录");
		new Label(parent, SWT.NONE);
		
		Button button = new Button(parent, SWT.NONE);
		button.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
			}
		});
		button.setLayoutData(new GridData(SWT.CENTER, SWT.CENTER, false, false, 1, 1));
		button.setText("开始");
		new Label(parent, SWT.NONE);
		
	}


	@Override
	public void setFocus() {
		// TODO Auto-generated method stub
		
	}
}
