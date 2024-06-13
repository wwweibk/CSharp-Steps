using Kompas6API5;

using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Kompas6Constants;
using KAPITypes;

namespace Steps.NET
{
	public class ksContrForm : System.Windows.Forms.Form
	{
		#region Designer declarations
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.Windows.Forms.OpenFileDialog openFileDialog;
		private System.Windows.Forms.Button btnGraphicLoad;
		private System.Windows.Forms.Button btnGetActive;
		private System.Windows.Forms.Button btnLoadFile;
		private System.Windows.Forms.Button btnNewFile;
		private System.Windows.Forms.Button btnLoadLibrary;
		private System.Windows.Forms.Button btnRunLibraryCommand;
		private System.Windows.Forms.Button btnSaveFile;
		private System.Windows.Forms.Button btnUnloadLibrary;
		private System.Windows.Forms.Button btnUnloadGraphic;
		private System.Windows.Forms.Button btnCloseFile;
		private System.Windows.Forms.Button btnQuitWOUnload;
		private System.Windows.Forms.Button btnQuitUnload;
		private System.Windows.Forms.Button btn��������������������;
		private System.Windows.Forms.Button btn����������������;
		private System.ComponentModel.Container components = null;
		#endregion

		#region Custom declarations
		private KompasObject kompas;
		int libraryId;					// HANDLE ����������� ����������
		#endregion

		#region Instance etc...
		private ksContrForm()
		{
			InitializeComponent();

		}


		private static ksContrForm instance;
		public static ksContrForm Instance
		{
			get
			{
				if (instance == null)
					instance = new ksContrForm();
				return instance;

			}
		}

		protected override void Dispose(bool disposing)
		{
			if(disposing)
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.btnGraphicLoad = new System.Windows.Forms.Button();
			this.btnGetActive = new System.Windows.Forms.Button();
			this.btnLoadFile = new System.Windows.Forms.Button();
			this.btnNewFile = new System.Windows.Forms.Button();
			this.btnLoadLibrary = new System.Windows.Forms.Button();
			this.btnRunLibraryCommand = new System.Windows.Forms.Button();
			this.btnSaveFile = new System.Windows.Forms.Button();
			this.btnUnloadLibrary = new System.Windows.Forms.Button();
			this.btnUnloadGraphic = new System.Windows.Forms.Button();
			this.btnCloseFile = new System.Windows.Forms.Button();
			this.btnQuitWOUnload = new System.Windows.Forms.Button();
			this.btnQuitUnload = new System.Windows.Forms.Button();
			this.btn�������������������� = new System.Windows.Forms.Button();
			this.btn���������������� = new System.Windows.Forms.Button();
			this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
			this.SuspendLayout();
			// 
			// btnGraphicLoad
			// 
			this.btnGraphicLoad.Location = new System.Drawing.Point(8, 8);
			this.btnGraphicLoad.Name = "btnGraphicLoad";
			this.btnGraphicLoad.Size = new System.Drawing.Size(216, 23);
			this.btnGraphicLoad.TabIndex = 0;
			this.btnGraphicLoad.Text = "Graphic Load";
			this.btnGraphicLoad.Click += new System.EventHandler(this.btnGraphicLoad_Click);
			// 
			// btnGetActive
			// 
			this.btnGetActive.Location = new System.Drawing.Point(8, 40);
			this.btnGetActive.Name = "btnGetActive";
			this.btnGetActive.Size = new System.Drawing.Size(216, 23);
			this.btnGetActive.TabIndex = 1;
			this.btnGetActive.Text = "Get Active";
			this.btnGetActive.Click += new System.EventHandler(this.btnGetActive_Click);
			// 
			// btnLoadFile
			// 
			this.btnLoadFile.Location = new System.Drawing.Point(9, 72);
			this.btnLoadFile.Name = "btnLoadFile";
			this.btnLoadFile.Size = new System.Drawing.Size(103, 23);
			this.btnLoadFile.TabIndex = 2;
			this.btnLoadFile.Text = "Load File";
			this.btnLoadFile.Click += new System.EventHandler(this.btnLoadFile_Click);
			// 
			// btnNewFile
			// 
			this.btnNewFile.Location = new System.Drawing.Point(120, 72);
			this.btnNewFile.Name = "btnNewFile";
			this.btnNewFile.Size = new System.Drawing.Size(104, 23);
			this.btnNewFile.TabIndex = 3;
			this.btnNewFile.Text = "New File";
			this.btnNewFile.Click += new System.EventHandler(this.btnNewFile_Click);
			// 
			// btnLoadLibrary
			// 
			this.btnLoadLibrary.Location = new System.Drawing.Point(9, 104);
			this.btnLoadLibrary.Name = "btnLoadLibrary";
			this.btnLoadLibrary.Size = new System.Drawing.Size(215, 23);
			this.btnLoadLibrary.TabIndex = 4;
			this.btnLoadLibrary.Text = "Load Library";
			this.btnLoadLibrary.Click += new System.EventHandler(this.btnLoadLibrary_Click);
			// 
			// btnRunLibraryCommand
			// 
			this.btnRunLibraryCommand.Location = new System.Drawing.Point(9, 136);
			this.btnRunLibraryCommand.Name = "btnRunLibraryCommand";
			this.btnRunLibraryCommand.Size = new System.Drawing.Size(215, 23);
			this.btnRunLibraryCommand.TabIndex = 5;
			this.btnRunLibraryCommand.Text = "Run Library Command";
			this.btnRunLibraryCommand.Click += new System.EventHandler(this.btnRunLibraryCommand_Click);
			// 
			// btnSaveFile
			// 
			this.btnSaveFile.Location = new System.Drawing.Point(9, 200);
			this.btnSaveFile.Name = "btnSaveFile";
			this.btnSaveFile.Size = new System.Drawing.Size(215, 23);
			this.btnSaveFile.TabIndex = 7;
			this.btnSaveFile.Text = "Save File";
			this.btnSaveFile.Click += new System.EventHandler(this.btnSaveFile_Click);
			// 
			// btnUnloadLibrary
			// 
			this.btnUnloadLibrary.Location = new System.Drawing.Point(9, 168);
			this.btnUnloadLibrary.Name = "btnUnloadLibrary";
			this.btnUnloadLibrary.Size = new System.Drawing.Size(215, 23);
			this.btnUnloadLibrary.TabIndex = 6;
			this.btnUnloadLibrary.Text = "Unload Library";
			this.btnUnloadLibrary.Click += new System.EventHandler(this.btnUnloadLibrary_Click);
			// 
			// btnUnloadGraphic
			// 
			this.btnUnloadGraphic.Location = new System.Drawing.Point(8, 264);
			this.btnUnloadGraphic.Name = "btnUnloadGraphic";
			this.btnUnloadGraphic.Size = new System.Drawing.Size(216, 23);
			this.btnUnloadGraphic.TabIndex = 9;
			this.btnUnloadGraphic.Text = "Unload Graphic";
			this.btnUnloadGraphic.Click += new System.EventHandler(this.btnUnloadGraphic_Click);
			// 
			// btnCloseFile
			// 
			this.btnCloseFile.Location = new System.Drawing.Point(8, 232);
			this.btnCloseFile.Name = "btnCloseFile";
			this.btnCloseFile.Size = new System.Drawing.Size(216, 23);
			this.btnCloseFile.TabIndex = 8;
			this.btnCloseFile.Text = "Close File";
			this.btnCloseFile.Click += new System.EventHandler(this.btnCloseFile_Click);
			// 
			// btnQuitWOUnload
			// 
			this.btnQuitWOUnload.Location = new System.Drawing.Point(120, 296);
			this.btnQuitWOUnload.Name = "btnQuitWOUnload";
			this.btnQuitWOUnload.Size = new System.Drawing.Size(104, 23);
			this.btnQuitWOUnload.TabIndex = 11;
			this.btnQuitWOUnload.Text = "Quit W/O Unload";
			this.btnQuitWOUnload.Click += new System.EventHandler(this.btnQuitWOUnload_Click);
			// 
			// btnQuitUnload
			// 
			this.btnQuitUnload.Location = new System.Drawing.Point(8, 296);
			this.btnQuitUnload.Name = "btnQuitUnload";
			this.btnQuitUnload.Size = new System.Drawing.Size(104, 23);
			this.btnQuitUnload.TabIndex = 10;
			this.btnQuitUnload.Text = "Quit && Unload";
			this.btnQuitUnload.Click += new System.EventHandler(this.btnQuitUnload_Click);
			// 
			// btn��������������������
			// 
			this.btn��������������������.Location = new System.Drawing.Point(8, 360);
			this.btn��������������������.Name = "btn��������������������";
			this.btn��������������������.Size = new System.Drawing.Size(216, 23);
			this.btn��������������������.TabIndex = 13;
			this.btn��������������������.Text = "���������� ����������";
			this.btn��������������������.Click += new System.EventHandler(this.btn��������������������_Click);
			// 
			// btn����������������
			// 
			this.btn����������������.Location = new System.Drawing.Point(8, 328);
			this.btn����������������.Name = "btn����������������";
			this.btn����������������.Size = new System.Drawing.Size(216, 23);
			this.btn����������������.TabIndex = 12;
			this.btn����������������.Text = "��������� �������";
			this.btn����������������.Click += new System.EventHandler(this.btn����������������_Click);
			// 
			// ksContrForm
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(234, 391);
			this.Controls.Add(this.btn��������������������);
			this.Controls.Add(this.btn����������������);
			this.Controls.Add(this.btnQuitWOUnload);
			this.Controls.Add(this.btnQuitUnload);
			this.Controls.Add(this.btnUnloadGraphic);
			this.Controls.Add(this.btnCloseFile);
			this.Controls.Add(this.btnSaveFile);
			this.Controls.Add(this.btnUnloadLibrary);
			this.Controls.Add(this.btnRunLibraryCommand);
			this.Controls.Add(this.btnLoadLibrary);
			this.Controls.Add(this.btnNewFile);
			this.Controls.Add(this.btnLoadFile);
			this.Controls.Add(this.btnGetActive);
			this.Controls.Add(this.btnGraphicLoad);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "ksContrForm";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Test Graphic Interface";
			this.ResumeLayout(false);

		}
		#endregion
		#endregion

		private void btnGraphicLoad_Click(object sender, System.EventArgs e)
		{
			if (kompas == null)
			{
				#if __LIGHT_VERSION__
					Type t = Type.GetTypeFromProgID("KOMPASLT.Application.5");
				#else
					Type t = Type.GetTypeFromProgID("KOMPAS.Application.5");
				#endif

				kompas = (KompasObject)Activator.CreateInstance(t);
			}

			if (kompas != null)
			{
				kompas.Visible = true;
				kompas.ActivateControllerAPI();
			}
		}

		private void btnGetActive_Click(object sender, System.EventArgs e)
		{
			if (kompas == null)
			{
				string progId = string.Empty;

				#if __LIGHT_VERSION__
					progId = "KOMPASLT.Application.5";
				#else
					progId = "KOMPAS.Application.5";
				#endif

				kompas = (KompasObject)Marshal.GetActiveObject(progId);
				if (kompas != null)
				{
					kompas.Visible = true;
					kompas.ActivateControllerAPI();
				}
				else
				{
					MessageBox.Show(this, "�� ������ �������� ������", "���������");
				}
			}
			else
			{
				MessageBox.Show(this, "������ ��� ��������", "���������");
			}
		}

		private void btnLoadFile_Click(object sender, System.EventArgs e)
		{
			if (kompas != null)
			{
				openFileDialog.Filter = "�������(*.cdw)|*.cdw|���������(*.frw)|*.frw|������(*.m3d)|*.m3d|������(*.a3d)|*.a3d|������������(*.spw)|*.spw";
				if (openFileDialog.ShowDialog(this) == DialogResult.OK)
				{
					// ������� �������� � �����
					// ������ �������� - ��� ������������ �����
					// ������ �������� ��������� �� ������������� ������ ������� "���� �������. ���������?" ��� �������� �����
					// ������ �������� - ��������� �� IDispatch, �� �������� ������ �������� ����������� �� ��������� ������ ���������
					// �-�� ���������� HANDLE ��������� ���������

					int type = kompas.ksGetDocumentTypeByName(openFileDialog.FileName);
					ksDocument3D doc3D;
					ksDocument2D doc2D;
					ksSpcDocument docSpc;
					ksDocumentTxt docTxt;
					switch (type) 
					{
						case (int)DocType.lt_DocPart3D:			//3d ���������
						case (int)DocType.lt_DocAssemble3D:
							doc3D = (ksDocument3D)kompas.Document3D();
							if (doc3D != null)
								doc3D.Open(openFileDialog.FileName, false);
							break;
						case (int)DocType.lt_DocSheetStandart :	//2d ���������
						case (int)DocType.lt_DocFragment:
							doc2D = (ksDocument2D)kompas.Document2D();
							if (doc2D != null)
								doc2D.ksOpenDocument(openFileDialog.FileName, false);
							break;
						case (int)DocType.lt_DocSpc:				//������������
							docSpc = (ksSpcDocument)kompas.SpcDocument();
							if (docSpc != null)
								docSpc.ksOpenDocument(openFileDialog.FileName, 0);
							break;
						case (int)DocType.lt_DocTxtStandart:		//��������� ���������
							docTxt = (ksDocumentTxt)kompas.DocumentTxt();
							if (docTxt != null)
								docTxt.ksOpenDocument(openFileDialog.FileName, 0);
							break;
					}
					int err = kompas.ksReturnResult();
					if (err != 0)
						kompas.ksResultNULL();
				}
			}
			else
			{
				MessageBox.Show(this, "������ �� ��������", "���������");
			}
		}

		private void btnNewFile_Click(object sender, System.EventArgs e)
		{
			if (kompas != null) 
			{
				// ������� ����� ��������
				// ������ �������� - ��� ������������ �����
				//  0 - ���� �������
				//  1 - ��������
				//  2 - ��������� ��������
				//  3 - ������������
				//  4 - 3D-������
				// ������ �������� ��������� �� ������������� ������ ������� "���� �������. ���������?" ��� �������� �����
				// ������ �������� - ��������� �� IDispatch, �� �������� ������ �������� ����������� �� �������� ������ ���������
				// �-�� ���������� HANDLE ��������� ���������
				ksDocument2D doc = (ksDocument2D)kompas.Document2D();    
				if (doc != null) 
				{
					ksDocumentParam docPar = (ksDocumentParam)kompas.GetParamStruct((short)StructType2DEnum.ko_DocumentParam);
					if (docPar != null) 
					{
						docPar.Init();
						docPar.type = (int)DocType.lt_DocSheetStandart;
						doc.ksCreateDocument(docPar); 
					}
				}
			}
			else
			{
				MessageBox.Show(this, "������ �� ��������", "���������");
			}
		}

		private void btnLoadLibrary_Click(object sender, System.EventArgs e)
		{
			if (kompas != null)
			{
				openFileDialog.Filter = "����������(*.rtw)|*.rtw";
				if (openFileDialog.ShowDialog(this) == DialogResult.OK)
				{
					// ��������� ����������
					// �-�� ���������� HANDLE ����������� ����������
					libraryId = kompas.ksAttachKompasLibrary(openFileDialog.FileName);
				}
			}
			else
			{
				MessageBox.Show(this, "������ �� ��������", "���������");
			}
		}

		private void btnRunLibraryCommand_Click(object sender, System.EventArgs e)
		{
			if (kompas != null) 
			{
				if (libraryId == 0)
					btnLoadLibrary_Click(null, null);
				// ��������� ������� � ����������� ����������
				// ����� �������� - HANDLE ����������
				// ������ �������� - ������������� ����������� �������
				kompas.ksExecuteKompasLibraryCommand(libraryId, 1);
			}
		}

		private void btnUnloadLibrary_Click(object sender, System.EventArgs e)
		{
			if (kompas != null && libraryId != 0) 
			{
				kompas.ksDetachKompasLibrary(libraryId);
				libraryId = 0;
			}
		}

		private void btnSaveFile_Click(object sender, System.EventArgs e)
		{
			if (kompas != null) 
			{
				ksDocument2D doc = (ksDocument2D)kompas.ActiveDocument2D();
				if (doc != null) 
					doc.ksSaveDocument(null);  
			}
			else
			{
				MessageBox.Show(this, "������ �� ��������", "���������");
			}
		}

		private void btnCloseFile_Click(object sender, System.EventArgs e)
		{
			if (kompas != null) 
			{
				ksDocument2D doc = (ksDocument2D)kompas.ActiveDocument2D();
				if (doc != null) 
					doc.ksCloseDocument(); 
			}
			else
			{
				MessageBox.Show(this, "������ �� ��������", "���������");
			}
		}

		private void btnUnloadGraphic_Click(object sender, System.EventArgs e)
		{
			if (kompas != null) 
			{
				kompas.Quit(); 
				Marshal.ReleaseComObject(kompas); 
			}
		}

		private void btnQuitUnload_Click(object sender, System.EventArgs e)
		{
			if (kompas != null) 
			{  
				kompas.Quit();  
				Marshal.ReleaseComObject(kompas); 
			}
			Close();
		}

		private void btnQuitWOUnload_Click(object sender, System.EventArgs e)
		{
			if (kompas != null) 
			{
				// ������������� ��������� ����� � ������
				Marshal.ReleaseComObject(kompas); 
			}
			Close(); // ��������� ������
		}

		private void btn����������������_Click(object sender, System.EventArgs e)
		{
			if (kompas != null) 
			{
				ksDocument2D doc = (ksDocument2D)kompas.ActiveDocument2D();
				if (doc != null) 
				{
					doc.ksCircle(50, 50, 20, 1);
					doc.ksCircle(50, 50, 50, 2);
					kompas.ksMessage("������");
				}
			}	
			else
			{
				MessageBox.Show(this, "������ �� ��������", "���������");
			}
		}

		private void btn��������������������_Click(object sender, System.EventArgs e)
		{
			try
			{
				if (kompas != null)
				{
					kompas.Visible = !kompas.Visible; 
				}
				else
					MessageBox.Show(this, "������ �� ��������", "���������");
			}
			catch (Exception ex)
			{
				ex.ToString();
				MessageBox.Show(this, "������, ����� ��������� � ��������", "���������");
				Marshal.ReleaseComObject(kompas);
				kompas = null;
			}
		}
	}
}
