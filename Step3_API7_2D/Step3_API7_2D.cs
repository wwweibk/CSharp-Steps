using Kompas6API5;
using KompasAPI7;
using Kompas6Constants;
using KAPITypes;
using System;
using System.Resources;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Windows.Forms;


namespace Step3_API7_2D
{
	public class SamplesSymbols
	{
    private KompasObject      kompas;
    private IApplication      appl;         // ��������� ����������
    private IKompasDocument2D doc;          // ��������� ��������� 2D � API7 
    private ksDocument2D      doc2D;        // ��������� ��������� 2D � API5
    private ResourceManager   resMng = new ResourceManager( typeof(SamplesSymbols) );       // �������� ��������


    //-------------------------------------------------------------------------------
    // ��������� ������ �� �������
    // ---
    string LoadString( string name )
    {
      return resMng.GetString( name );
    }
    
    //-------------------------------------------------------------------------------
    // ��� ����������
    // ---
    [return: MarshalAs(UnmanagedType.BStr)] public string GetLibraryName()
    {
      return LoadString( "IDS_LIBNAME" );
    }

    //-------------------------------------------------------------------------------
    // ���� ����������
    // ---
    [return: MarshalAs(UnmanagedType.BStr)] public string ExternalMenuItem(short number, ref short itemType, ref short command)
    {
      string result = string.Empty;
      itemType = 1; // "MENUITEM"
      switch (number)
      {
        case 1:
          result = "�������� � �������������� ������� ����� �������";
          command = 1;
        break;

        case 2:
          result = "�������� � �������������� ����� ����������";
          command = 2;
        break;

        case 3:
          result = "�������� � �������������� ����� ���������";
          command = 3;
        break;

        case 4:
          result = "�������� � �������������� ����� ���������";
          command = 4;
        break;

        case 5:
          result = "�������� � �������������� ����������� �������������";
          command = 5;
        break;

        case 6:
          result = "�������� � �������������� ����������� ����";
          command = 6;
        break;

        case 7:
          result = "�������� � �������������� ����� �������/�������";
          command = 7;
        break;

        case 8:
          result = "�������� � �������������� ������� ����������� �������";
          command = 8;
        break;

        case 9:
          result = "�������� � �������������� ������� �����";
          command = 9;
        break;

        case 10:
          result = "��������� �� ������� �������� ����";
          command = 10;
        break;

        case 11:
          command = -1;
          itemType = 3; // "ENDMENU"
        break;
      }
      return result;
    }


    //-------------------------------------------------------------------------------
    // �������� ������� ����������
    // ---
    public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
    {
      kompas = (KompasObject)kompas_;

      if ( kompas == null )
        return;
      
      // �������� ��������� ����������
      appl = (IApplication)kompas.ksGetApplication7();

      if ( appl == null )
        return;

      // �������� ��������� ��������� ��������� 2D � API7
      doc = (IKompasDocument2D)appl.ActiveDocument;

      if ( doc == null )
        return;

      // �������� ��������� ��������� ��������� 2D � API5
      doc2D = (ksDocument2D)kompas.ActiveDocument2D();

      if ( doc2D == null )
        return;

      switch ( command )
      {
        case 1: LeaderWork();       break;  // �������� � �������������� ������� ����� �������
        case 2: MarkLeaderWork();   break;  // �������� � �������������� ����� ����������
        case 3: ChangeLeaderWork(); break;  // �������� � �������������� ����� ���������
        case 4: BrandLeaderWork();  break;  // �������� � �������������� ����� ���������
        case 5: RoughWork();        break;  // �������� � �������������� ����������� �������������
        case 6: BaseWork();         break;  // �������� � �������������� ����������� ����
        case 7: CutLineWork();      break;  // �������� � �������������� ����� �������/�������
        case 8: ViewPointerWork();  break;  // �������� � �������������� ������� ����������� �������
        case 9: ToleranceWork();    break;  // �������� � �������������� ������� �����
        case 10: ObjectsNavigation(); break;  // ��������� �� ������� �������� ����
      }
    }

    #region ��������������� �������
    //-------------------------------------------------------------------------------
    //  �������� ��������� ����������� 2D
    // ---
    ISymbols2DContainer GetSymbols2DContainer()
    {
      if ( doc != null )
      {
        // ������� �������� ��� ������ � ������ � ������
        ViewsAndLayersManager viewsMng = doc.ViewsAndLayersManager;
    		
        if ( viewsMng != null )
        {
          // ������� ��������� �����
          IViews views = viewsMng.Views;
    			
          if ( views != null )
          {
            // �������� ��������� � ��������� ����
            IView view = views.ActiveView;

            if ( view != null )
              return (ISymbols2DContainer)view;
          }
        }
      }
      return null;
    }


    //-------------------------------------------------------------------------------
    // ������� ����� � �������� �������� ������
    // ---
    void SetTextSmallFont( ref IText txt, string str, double size )
    {
	    if ( txt != null )
	    {
		    // �������� ��������� ������ ������
		    ITextLine line = txt.Add();

		    if ( line != null )
		    {
			    // �������� ��������� ���������� ������
			    ITextItem item = line.Add();

			    if ( item != null )
			    {
				    // �������� ��������� ������
				    ITextFont font = (ITextFont)item;

				    if ( font != null )
					    // ������ ������ ������
					    font.Height = size;

				    // ������ �������� ������
				    item.Str = str;
				    // ��������� ���������
				    item.Update();
			    }
		    }
	    }
    }
    #endregion


    #region ������� ����� �������
    //-------------------------------------------------------------------------------
    // �������� ������� ����� �������
    // ---
    bool CreateLeader( ref ILeader leader )
    {
      bool res = false;

	    if ( leader != null )
	    {
		    // ����������� ����� - ������
		    leader.ShelfDirection = Kompas6Constants.ksShelfDirectionEnum.ksLSRight;

		    // �������� ��������� �����������
		    IBranchs branchs = (IBranchs)leader;

		    if ( branchs != null )
		    {
			    // ���������� ������ ����� ��� ����� ��������
			    branchs.X0 = 100;
			    branchs.Y0 = 150;
			    // �������� ������������� �����������
			    branchs.AddBranchByPoint( -1, 60, 120 );
			    branchs.AddBranchByPoint( -1, 65, 105 );
		    }

		    // �������� ��������� ������ ��� ������
		    IText txtOnSh = leader.TextOnShelf;

		    if ( txtOnSh != null )
			    // �������� �����
			    txtOnSh.Str = "1";

		    // �������� ��������� ������ ��� ������
		    IText txtUnderSh = leader.TextUnderShelf;

		    if ( txtUnderSh != null )
			    // �������� �����
			    txtUnderSh.Str = "2";

		    // �������� ��������� ������ ��� ������
		    IText txtOnBr = leader.TextOnBranch;
    		
		    if ( txtOnBr != null )
			    // �������� �����
			    txtOnBr.Str = "3";

		    // �������� ��������� ������ ��� ������
		    IText txtUnderBr = leader.TextUnderBranch;
    		
		    if ( txtUnderBr != null )
			    // �������� �����
			    txtUnderBr.Str = "4";

		    // �������� ��������� ������ �� ������
		    IText txtAfterSh = leader.TextAfterShelf;
    		
		    if ( txtAfterSh != null )
			    // �������� �����
			    txtAfterSh.Str = "5";

		    IBaseLeader baseLeader = (IBaseLeader)leader;

		    if ( baseLeader != null )
			    // ��������� ���������
			    res = baseLeader.Update();
	    }
      return res;
    }


    //-------------------------------------------------------------------------------
    // ������ ���������� ������
    // ---
    void AddTextItem( ref ITextLine line, string str, Kompas6Constants.ksTextItemEnum type )
    {
	    if ( line != null )
	    {
		    // �������� ��������� ���������� ������
		    ITextItem item = line.Add();
    		
		    if ( item != null )
		    {
			    // ������ ��������� ��������
			    item.Str = str;
			    // ������ ���
			    item.ItemType = type;
			    // ��������� ���������
			    item.Update();
		    }
	    }
    }


    //-------------------------------------------------------------------------------
    // �������������� ����� �������
    // ---
    void EditLeader( ref ILeader leader )
    {
	    if ( leader != null )
	    {
		    // ��� ������ - ���� ����������
		    leader.SignType = Kompas6Constants.ksLeaderSignEnum.ksLGlueSign;
		    // �������� ������� ���������� �� �������
		    leader.Arround = true;

		    // �������� ��������� �����������
		    IBranchs branchs = (IBranchs)leader;
    		
		    if ( branchs != null )
			    // �������� ������������� �����������
			    branchs.AddBranchByPoint( -1, 140, 120 );

		    // ������ ���������� ����������� - �� ����� �����
		    leader.set_BranchBegin( 2, false );

		    // �������� ����� ��� ������
		    IText txt = leader.TextOnShelf;

		    if ( txt != null )
		    {
			    txt.Str = "";

			    ITextLine line = txt.Add();

			    if ( line != null )
			    {
				    // �������� ������
				    AddTextItem( ref line, "1", Kompas6Constants.ksTextItemEnum.ksTItString );
				    // �������� ��������� �����
				    AddTextItem( ref line, "1", Kompas6Constants.ksTextItemEnum.ksTItNumerator );
				    // �������� ������������ �����
				    AddTextItem( ref line, "1", Kompas6Constants.ksTextItemEnum.ksTItDenominator );
				    // ��������� �����
				    AddTextItem( ref line, "1", Kompas6Constants.ksTextItemEnum.ksTItFractionEnd );
			    }
		    }

		    IBaseLeader baseLeader = (IBaseLeader)leader;
    		
		    if ( baseLeader != null )
			    // ��������� ���������
			    baseLeader.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // �������� � �������������� ������� ����� �������
    // ---
    void LeaderWork()
    {
	    // �������� ��������� �������� �����������
	    ISymbols2DContainer symbCont = GetSymbols2DContainer();

	    if ( symbCont != null )
	    {
		    // �������� ��������� ����� �������
		    ILeaders leadersCol = symbCont.Leaders;

		    if ( leadersCol != null )
		    {
			    // �������� ������� ����� �������
			    ILeader leader = (ILeader)leadersCol.Add( Kompas6Constants.DrawingObjectTypeEnum.
                                           ksDrLeader );

			    if ( leader != null && doc2D != null )
			    {
				    // ������� ����� �������
				    bool create = CreateLeader( ref leader );

				    // �������� �������� �������
				    IBaseLeader bLeader = (IBaseLeader)leader;
				    int refr = 0;

				    if ( bLeader != null )
					    refr = bLeader.Reference; 
    				
				    // ���������� ��������� ������
            doc2D.ksLightObj( refr, 1/*�������� ���������*/ );

            if ( create && kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
				    {		
					    // �������� ������ �� ��������� �� ���������
					    ILeader lead = (ILeader)leadersCol.get_Leader( refr );				
					    // ������������� ����� �������
					    EditLeader( ref lead );
				    }
            doc2D.ksLightObj( refr, 0/*��������� ���������*/ );
			    }
		    }
	    }
    }
    #endregion


    #region ���� ����������
    //-------------------------------------------------------------------------------
    // �������� ����� ����������
    // ---
    bool CreateMarkLeader( ref IMarkLeader markLeader )
    {
      bool res = false;

	    if ( markLeader != null )
	    {
		    // �������� ��������� �����������
		    IBranchs branchs = (IBranchs)markLeader;
    		
		    if ( branchs != null )
		    {
			    // ����� �������� 
			    branchs.X0 = 100;
			    branchs.Y0 = 190;
			    // �������� ������������� �����������
			    branchs.AddBranchByPoint( -1, 60, 120 );
		    }

		    // �������� ��������� ������ �����������
		    IText des = markLeader.Designation;

		    if ( des != null )
			    // �������� �����
			    des.Str = LoadString( "IDS_MARK" );

		    // �������� ��������� ������ ��� ������
		    IText textOnBranch = markLeader.TextOnBranch;

		    if ( textOnBranch != null )
			    // �������� �����
			    textOnBranch.Str = "2";

		    // �������� ��������� ������ ��� ������
		    IText textUnderBranch = markLeader.TextUnderBranch;
    		
		    if ( textUnderBranch != null )
			    // �������� �����
			    textUnderBranch.Str = "3";

		    // �������� ������� ��������� ����� �������
		    IBaseLeader baseLeader = (IBaseLeader)markLeader;

		    if ( baseLeader != null )
			    // ��������� ���������
			    res = baseLeader.Update();
	    }
      return res;
    }


    //-------------------------------------------------------------------------------
    // �������������� ����� ����������
    // ---
    void EditMarkLeader( ref IMarkLeader markLeader )
    {
	    if ( markLeader != null )
	    {
		    // �������� ��������� �����������
		    IBranchs branchs = (IBranchs)markLeader;

		    if ( branchs != null )
			    // �������� ������������� �����������
			    branchs.AddBranchByPoint( -1, 70, 110 );

		    // �������� ������� ��������� ����� �������
		    IBaseLeader baseLeader = (IBaseLeader)markLeader;

		    if ( baseLeader != null )
		    {
			    // ��� �������
			    baseLeader.ArrowType = Kompas6Constants.ksArrowEnum.ksLeaderArrow;
			    // ��������� ���������
			    baseLeader.Update();
		    }
	    }
    }


    //-------------------------------------------------------------------------------
    // �������� � �������������� ����� ����������
    // ---
    void MarkLeaderWork()
    {
	    // �������� ��������� �������� �����������
	    ISymbols2DContainer symbCont = GetSymbols2DContainer();
    	
	    if ( symbCont != null )
	    {
		    // �������� ��������� ����� �������
		    ILeaders leadersCol = symbCont.Leaders;
    		
		    if ( leadersCol != null )
		    {
			    // �������� ���� ����������
			    IMarkLeader mLeader = (IMarkLeader)leadersCol.Add( Kompas6Constants.
                                 DrawingObjectTypeEnum.ksDrMarkerLeader);
    			
			    if ( mLeader != null && doc2D != null )
			    {
				    // ������� ���� ����������
				    bool create = CreateMarkLeader( ref mLeader );
    				
				    // �������� �������� �������
				    IBaseLeader bLeader = (IBaseLeader)mLeader;
				    int refr = 0;
    				
				    if ( bLeader != null )
					    refr = bLeader.Reference;
    				
				    // ���������� ��������� ������
            doc2D.ksLightObj( refr, 1/*�������� ���������*/ );
    				
            if ( create && kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
				    {
					    // �������� ������ �� ��������� �� ���������
					    IMarkLeader markLead = (IMarkLeader)leadersCol.get_Leader( refr );				
					    // ������������� ���� ����������
					    EditMarkLeader( ref markLead );
				    }
            doc2D.ksLightObj( refr, 0/*��������� ���������*/ );
			    }
		    }
	    }
    }
    #endregion


    #region ���� ���������
    //-------------------------------------------------------------------------------
    // �������� ����� ���������
    // ---
    bool CreateChangeLeader( ref IChangeLeader changeLeader )
    {
      bool res = false;

	    if ( changeLeader != null )
	    {
		    // ��� ������ - �������
		    changeLeader.SignType = Kompas6Constants.ksChangeLeaderSignEnum.ksCLSSquare;

		    // �������� ��������� �����������
		    IBranchs branchs = (IBranchs)changeLeader;

		    if ( branchs != null )
		    {
			    // ����� ��������
			    branchs.X0 = 70;
			    branchs.Y0 = 150;
			    // �������� �����������
			    branchs.AddBranchByPoint( -1, 40, 130 );
		    }

		    // �������� ��������� ������ �����������
		    IText des = changeLeader.Designation;
    		
		    if ( des != null )
			    // �������� �����
			    des.Str = "1";

		    // �������� ������� ��������� ����� �������
		    IBaseLeader baseLeader = (IBaseLeader)changeLeader;

		    if ( baseLeader != null )
			    // ��������� ���������
			    res = baseLeader.Update();
	    }
      return res;
    }


    //-------------------------------------------------------------------------------
    // �������������� ����� ���������
    // ---
    void EditChangeLeader( ref IChangeLeader changeLeader )
    {
	    if ( changeLeader != null )
	    {
		    // ��� ������ - ����
		    changeLeader.SignType = Kompas6Constants.ksChangeLeaderSignEnum.ksCLSCircle;
		    // ������ ����� ������� ���������
		    changeLeader.FullLeaderLength = false;
		    // ����� �������
		    changeLeader.LeaderLength = 5;

		    // �������� ������� ��������� ����� �������
		    IBaseLeader baseLeader = (IBaseLeader)changeLeader;
    		
		    if ( baseLeader != null )
			    // ��������� ���������
			    baseLeader.Update();
	    }
    }

    //-------------------------------------------------------------------------------
    // �������� � �������������� ����� ���������
    // ---
    void ChangeLeaderWork()
    {
	    // �������� ��������� �������� �����������
	    ISymbols2DContainer symbCont = GetSymbols2DContainer();
    	
	    if ( symbCont != null )
	    {
		    // �������� ��������� ����� �������
		    ILeaders leadersCol = symbCont.Leaders;
    		
		    if ( leadersCol != null )
		    {
			    // �������� ���� ���������
			    IChangeLeader chLeader = (IChangeLeader)leadersCol.Add( Kompas6Constants.
                                    DrawingObjectTypeEnum.ksDrChangeLeader );
    			
			    if ( chLeader != null && doc2D != null )
			    {
				    // ������� ���� ���������
				    bool create = CreateChangeLeader( ref chLeader );
    				
				    // �������� �������� �������
				    IBaseLeader bLeader = (IBaseLeader)chLeader;
				    int refr = 0;
    				
				    if ( bLeader != null )
					    refr = bLeader.Reference;
    				
				    // ���������� ��������� ������
            doc2D.ksLightObj( refr, 1/*�������� ���������*/ );
    				
            if ( create && kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
				    {
					    // �������� ������ �� ��������� �� ���������
					    IChangeLeader changeLead = (IChangeLeader)leadersCol.get_Leader( refr );				
					    // ������������� ���� ���������
					    EditChangeLeader( ref changeLead );
				    }
            doc2D.ksLightObj( refr, 0/*��������� ���������*/ );
			    }
		    }
	    }
    }
    #endregion


    #region ���� ���������
    //-------------------------------------------------------------------------------
    // �������� ����� ���������
    // ---
    bool CreateBrandLeader( ref IBrandLeader brandLeader )
    {
      bool res = false;

	    if ( brandLeader != null )
	    {
		    // �����������
		    brandLeader.Direction = false;

		    // �������� ��������� ������ �����������
		    IText des = brandLeader.Designation;

		    if ( des != null )
			    // �������� �����
			    des.Str = LoadString( "IDS_MARK2" );

		    // �������� ��������� �����������
		    IBranchs branchs = (IBranchs)brandLeader;

		    if ( branchs != null )
		    {
			    // ����� ��������
			    branchs.X0 = 100;
			    branchs.Y0 = 150;
			    // �������� �����������
			    branchs.AddBranchByPoint( -1, 60, 110 );
		    }

		    // �������� ������� ��������� ����� �������
		    IBaseLeader baseLeader = (IBaseLeader)brandLeader;
    		
		    if ( baseLeader != null )
			    // ��������� ���������
			    res = baseLeader.Update();
	    }
      return res;
    }


    //-------------------------------------------------------------------------------
    // �������������� ����� ���������
    // ---
    void EditBrandLeader( ref IBrandLeader brandLeader )
    {
	    if ( brandLeader != null )
	    {
		    // �����������
		    brandLeader.Direction = true;

		    // �������� ��������� ������ ��� ������
		    IText textOn = brandLeader.TextOnBranch;
    		
		    if ( textOn != null )
			    // �������� �����
			    textOn.Str = "2";
    		
		    // �������� ��������� ������ ��� ������
		    IText textUnder = brandLeader.TextUnderBranch;
    		
		    if ( textUnder != null )
			    // �������� �����
			    textUnder.Str = "3";

		    // �������� ��������� ������ �����������
		    IText des = brandLeader.Designation;
    		
		    if ( des != null )
			    // �������� �����
			    des.Str = LoadString( "IDS_MARK3" );

		    // �������� ������� ��������� ����� �������
		    IBaseLeader baseLeader = (IBaseLeader)brandLeader;
    		
		    if ( baseLeader != null )
			    // ��������� ���������
			    baseLeader.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // �������� � �������������� ����� ���������
    // ---
    void BrandLeaderWork()
    {
	    // �������� ��������� �������� �����������
	    ISymbols2DContainer symbCont = GetSymbols2DContainer();
    	
	    if ( symbCont != null )
	    {
		    // �������� ��������� ����� �������
		    ILeaders leadersCol = symbCont.Leaders;
    		
		    if ( leadersCol != null )
		    {
			    // �������� ���� ���������
			    IBrandLeader brLeader = (IBrandLeader)leadersCol.Add( Kompas6Constants.DrawingObjectTypeEnum.ksDrBrandLeader );
    			
			    if ( brLeader != null && doc2D != null )
			    {
				    // ������� ���� ���������
				    bool create = CreateBrandLeader( ref brLeader );
    				
				    // �������� �������� �������
				    IBaseLeader bLeader = (IBaseLeader)brLeader;
				    int refObj = 0;
    				
				    if ( bLeader != null )
					    refObj = bLeader.Reference;
    				
				    // ���������� ��������� ������
            doc2D.ksLightObj( refObj, 1/*�������� ���������*/ );
    				
            if ( create && kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
				    {
					    // �������� ������ �� ��������� �� ���������
					    IBrandLeader brandLead = (IBrandLeader)leadersCol.get_Leader( refObj );				
					    // ������������� ���� ���������
					    EditBrandLeader( ref brandLead );
				    }
            doc2D.ksLightObj( refObj, 0/*��������� ���������*/ );
			    }
		    }
	    }
    }
    #endregion


    #region ����������� �������������
    //-------------------------------------------------------------------------------
    // �������� ����������� �������������
    // ---
    bool CreateRough( ref IRough rough )
    {
	    bool res = false;

	    if ( rough != null && doc2D != null )
	    {
		    // ��������� ���������� ������� � �������
		    ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct( (int)Kompas6Constants.
                              StructType2DEnum.ko_RequestInfo );
		    info.Init();
		    info.commandsString = LoadString( "IDS_COMMAND1" );
        info.cursorId = ldefin2d.OCR_CATCH;
		    double x = 0, y = 0;

		    // ��������� ��������
		    SnapOptions sOpt = (SnapOptions)kompas.GetParamStruct( (int)Kompas6Constants.
                            StructType2DEnum.ko_SnapOptions );
        kompas.ksGetSysOptions( ldefin2d.SNAP_OPTIONS, sOpt );
		    sOpt.commonOpt = 0;
		    kompas.ksSetSysOptions( ldefin2d.SNAP_OPTIONS, sOpt );

		    // ������� ������� ������
        if ( doc2D.ksCursorEx(info, ref x, ref y, null, null) != 0 )
		    {
			    // ����� ������ �� ��������� �����������
          int pObj = doc2D.ksFindObj( x, y, doc2D.ksGetCursorLimit() );
			    // ���������� ������
          doc2D.ksLightObj( pObj, 1/*�������� ���������*/ );
			    // �������� ��������� ������������ ������� �� ���������
			    IDrawingObject baseObj = (IDrawingObject)kompas.TransferReference( pObj, doc2D.reference );
			    info.commandsString = LoadString( "IDS_COMMAND2" );
			    info.cursorId = 0;

			    // ���� �������� ���������, ��������
			    if ( sOpt.commonOpt == 0 )
			    {
				    sOpt.commonOpt = 1;
				    kompas.ksSetSysOptions( ldefin2d.SNAP_OPTIONS, sOpt );
			    }

			    // ������� ����� ��������
          if ( doc2D.ksCursorEx(info, ref x, ref y, null, null) != 0 )
			    {
				    if ( baseObj != null )
					    // ������� ������
					    rough.BaseObject = baseObj;

				    // ����� ��������
				    rough.BranchX0 = x;
				    rough.BranchY0 = y;
    				
				    IRoughParams roughPar = (IRoughParams)rough;

				    if ( roughPar != null )
				    {
					    roughPar.ShelfDirection = Kompas6Constants.ksShelfDirectionEnum.ksLSRight;
					    rough.Update();
					    // ��������� �� ������� ��������
					    roughPar.ProcessingByContour = true;
					    // ����� ����� �������
					    roughPar.LeaderLength = 20;
					    // ���� ������� ����� �������
					    roughPar.LeaderAngle = 45;

					    // �������� ��������� ������ ���������� �������������
					    IText txt1 = roughPar.RoughParamText;

					    if ( txt1 != null )
						    txt1.Str = "1";
    					
					    // �������� ��������� ������ ������� ��������� �����������
					    IText txt2 = roughPar.ProcessText;
    					
					    if ( txt2 != null )
						    txt2.Str = "2";

					    // �������� ��������� ������ ������� �����
					    IText txt3 = roughPar.BaseLengthText;
    					
					    if ( txt3 != null )
						    txt3.Str = "3";
				    }

				    // ��������� ���������
				    res = rough.Update();
			    }
			    else
            kompas.ksMessage( LoadString("IDS_NOPOINT") );

			    doc2D.ksLightObj( pObj, 0/*��������� ���������*/ );
		    }
		    else
          kompas.ksMessage( LoadString("IDS_NOOBJ") );

		    // ���� �������� ���������, ��������
		    if ( sOpt.commonOpt == 0 )
		    {
			    sOpt.commonOpt = 1;
			    kompas.ksSetSysOptions( ldefin2d.SNAP_OPTIONS, sOpt );
		    }
	    }
	    return res;
    }


    //-------------------------------------------------------------------------------
    // �������������� ����������� �������������
    // ---
    void EditRough( ref IRough rough )
    {
	    if ( rough != null )
	    {
		    // �������� ��������� ���������� �������������
		    IRoughParams roughPar = (IRoughParams)rough;

		    if ( roughPar != null )
		    {
			    // ������ �������
			    roughPar.ArrowType = Kompas6Constants.ksArrowEnum.ksWithoutArrow;
			    // ��������� �� ������� ���������
			    roughPar.ProcessingByContour = false;
		    }
		    // ��������� ���������
		    rough.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // �������� � �������������� ����������� �������������
    // ---
    void RoughWork()
    {
	    // �������� ��������� �������� �����������
	    ISymbols2DContainer symbCont = GetSymbols2DContainer();
    	
	    if ( symbCont != null )
	    {
		    // �������� ��������� ����������� �������������
		    IRoughs roughsCol = symbCont.Roughs;
    		
		    if ( roughsCol != null )
		    {
			    // �������� ����������� �������������
			    IRough newRough = roughsCol.Add();
    			
			    if ( newRough != null && doc2D != null )
			    {
				    // ������� ����������� �������������
				    if ( CreateRough(ref newRough) )
				    {			
					    // �������� �������� �������
					    int objRef = newRough.Reference;
					    // ���������� ��������� ������
              doc2D.ksLightObj( objRef, 1/*�������� ���������*/ );
    					
              if ( kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
					    {
						    // �������� ������ �� ��������� �� ���������
						    IRough rough = roughsCol.get_Rough( objRef );				
						    // ������������� ����������� �������������
						    EditRough( ref rough );
					    }
              doc2D.ksLightObj( objRef, 0/*��������� ���������*/ );
				    }
			    }
		    }
	    }
    }
    #endregion


    #region ����������� ����
    //-------------------------------------------------------------------------------
    // �������� ����������� ����
    // ---
    bool CreateBase( ref IBase pBase )
    {
	    bool res = false;

	    if ( pBase != null )
	    {
        // ��������� ���������� ������� � �������
        ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct( (int)Kompas6Constants.
                              StructType2DEnum.ko_RequestInfo );
        info.Init();
        info.commandsString = LoadString( "IDS_COMMAND1" );
        info.cursorId = ldefin2d.OCR_CATCH;
        double x = 0, y = 0;

        // ��������� ��������
        SnapOptions sOpt = (SnapOptions)kompas.GetParamStruct( (int)Kompas6Constants.
                            StructType2DEnum.ko_SnapOptions );
        kompas.ksGetSysOptions( ldefin2d.SNAP_OPTIONS, sOpt );
        sOpt.commonOpt = 0;
        kompas.ksSetSysOptions( ldefin2d.SNAP_OPTIONS, sOpt );
    		
		    // ������� ������� ������
        if ( doc2D.ksCursorEx(info, ref x, ref y, null, null) != 0 )
		    {
			    // ����� ������ �� ��������� �����������
          int pObj = doc2D.ksFindObj( x, y, doc2D.ksGetCursorLimit() );
			    // ���������� ������
          doc2D.ksLightObj( pObj, 1/*�������� ���������*/ );
			    // �������� ��������� ������������ ������� �� ���������
			    IDrawingObject baseObj = (IDrawingObject)kompas.TransferReference( pObj, doc2D.reference );
			    info.commandsString = LoadString( "IDS_COMMAND2" );
			    info.cursorId = 0;

			    // ���� �������� ���������, ��������
			    if ( sOpt.commonOpt == 0 )
			    {
				    sOpt.commonOpt = 1;
				    kompas.ksSetSysOptions( ldefin2d.SNAP_OPTIONS, sOpt );
			    }
    			
			    // ������� ��������� �����
          if ( doc2D.ksCursorEx(info, ref x, ref y, null, null) != 0 )
			    {
				    if ( baseObj != null )
					    // ������� ������
					    pBase.BaseObject = baseObj;

				    // ����� ��������� �����
				    pBase.X0 = x;
				    pBase.Y0 = y;
				    // ������ ��������� - ��������������� � �������
				    pBase.DrawType = true;
				    // ��������� ���������
				    res = pBase.Update();
			    }
			    else
            kompas.ksMessage( LoadString("IDS_NOPOINT") );

          doc2D.ksLightObj( pObj, 0/*��������� ���������*/ );
		    }
		    else
          kompas.ksMessage( LoadString("IDS_NOOBJ") );

		    // ���� �������� ���������, ��������
		    if ( sOpt.commonOpt == 0 )
		    {
			    sOpt.commonOpt = 1;
			    kompas.ksSetSysOptions( ldefin2d.SNAP_OPTIONS, sOpt );
		    }
	    }
	    return res;
    }


    //-------------------------------------------------------------------------------
    // �������������� ����������� ����
    // ---
    void EditBase( ref IBase pBase )
    {
	    if ( pBase != null )
	    {
		    // ��� ��������� - �����������
		    pBase.DrawType = false;
		    // ��������� ��������������
		    pBase.AutoSorted = false;
		    // �������� ����� �������
		    pBase.BranchX = pBase.X0 + 10;
		    pBase.BranchY = pBase.Y0 + 10;

		    // �������� ��������� ������ ����������� ����
		    IText txt = pBase.Text;

		    if ( txt != null )
			    // �������� �����
			    txt.Str = "B";

		    // ��������� ���������
		    pBase.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // �������� � �������������� ����������� ����
    // ---
    void BaseWork()
    {
	    // �������� ��������� �������� �����������
	    ISymbols2DContainer symbCont = GetSymbols2DContainer();
    	
	    if ( symbCont != null )
	    {
		    // �������� ��������� ����������� ����
		    IBases basesCol = symbCont.Bases;
    		
		    if ( basesCol != null )
		    {
			    // �������� ����������� ����
			    IBase newBase = basesCol.Add();
    			
			    if ( newBase != null )
			    {
				    // ������� ����������� ����
				    if ( CreateBase( ref newBase) )
				    {			
					    // �������� �������� �������
					    int objRef = newBase.Reference;
					    // ���������� ��������� ������
              doc2D.ksLightObj( objRef, 1/*�������� ���������*/ );
    					
              if ( kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
					    {
						    // �������� ������ �� ��������� �� ���������
						    IBase pBase = basesCol.get_Base( objRef );				
						    // ������������� ����������� ����
						    EditBase( ref pBase );
					    }
              doc2D.ksLightObj( objRef, 0/*��������� ���������*/ );
				    }
			    }
		    }
	    }
    }
    #endregion


    #region ����� �������/�������
    //-------------------------------------------------------------------------------
    // �������� ����� �������/�������
    // ---
    bool CreateCutLine( ref ICutLine cutLine )
    {
      bool res = false;
 
	    if ( cutLine != null )
	    {
		    // ���������� ���������� ������
		    cutLine.X1 = 80;
		    cutLine.Y1 = 165;
		    // ���������� ��������� ������
		    cutLine.X2 = 120;
		    cutLine.Y2 = 200;
		    // ������������ ������� - �����
		    cutLine.ArrowPos = true;
		    // ���������� ��������������� ������ - � ������ �������
		    cutLine.AdditionalTextPos = true;

		    // ������� ������ ����� ����� �������
		    double[] points = new double[4];
		    points[0] = 80;
		    points[1] = 165;
		    points[2] = 120;
		    points[3] = 200;
			  // ������ ������ ����� ����� �������
			  cutLine.Points = points;

		    // �������� ��������� ��������������� ������
		    IText adText = cutLine.AdditionalText;
    		
		    if ( adText != null )
			    SetTextSmallFont( ref adText, "(1)", 7 );

		    // ��������� ���������
		    res = cutLine.Update();
	    }
      return res;
    }


    //-------------------------------------------------------------------------------
    // �������������� ����� �������/�������
    // ---
    void EditCutLine( ref ICutLine cutLine )
    {
	    if ( cutLine != null )
	    {
		    // ������������ ������� - ������ �� ����������� �������
		    cutLine.ArrowPos = false;
		    // ������������ ��������������� ������ - � ������ �������
		    cutLine.AdditionalTextPos = false;
		    // ��������� ��������������
		    cutLine.AutoSorted = false;

		    // ������� ������ ����� ����� �������
		    double[] points = new double[6];
		    points[0] = 80;
		    points[1] = 165;
		    points[2] = 115;
		    points[3] = 165;
		    points[4] = 120;
		    points[5] = 200;
			  cutLine.Points = points;

		    // ��������� ���������
		    cutLine.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // �������� � �������������� ����� �������/�������
    // ---
    void CutLineWork()
    {
	    // �������� ��������� �������� �����������
	    ISymbols2DContainer symbCont = GetSymbols2DContainer();
    	
	    if ( symbCont != null )
	    {
		    // �������� ��������� ����� �������/�������
		    ICutLines cutCol = symbCont.CutLines;
    		
		    if ( cutCol != null && doc2D != null )
		    {
			    // �������� ����� �������/�������
			    ICutLine newCut = cutCol.Add();
    			
			    if ( newCut != null )
			    {
				    // ������� ����� �������/�������
				    bool create = CreateCutLine( ref newCut );
    				
				    // �������� �������� �������
				    int objRef = newCut.Reference;
				    // ���������� ��������� ������
            doc2D.ksLightObj( objRef, 1/*�������� ���������*/ );
    				
            if ( create && kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
				    {
					    // �������� ������ �� ��������� �� ���������
					    ICutLine cutLine = cutCol.get_CutLine( objRef );				
					    // ������������� ����� �������/�������
					    EditCutLine( ref cutLine );
				    }
            doc2D.ksLightObj( objRef, 0/*��������� ���������*/ );
			    }
		    }
	    }
    }
    #endregion


    #region ������� ����������� �������
    //-------------------------------------------------------------------------------
    // �������� ������� ����������� �������
    // ---
    bool CreateViewPointer( ref IViewPointer viewPointer )
    {
      bool res = false;

	    if ( viewPointer != null )
	    {
		    // ��������� ����� �������
		    viewPointer.X1 = 100;
		    viewPointer.Y1 = 150;
		    // �������� ����� �������
		    viewPointer.X2 = 120;
		    viewPointer.Y2 = 160;
		    // ��������� ���������
		    res = viewPointer.Update();
	    }
      return res;
    }


    //-------------------------------------------------------------------------------
    // �������������� ������� ����������� �������
    // ---
    void EditViewPointer( ref IViewPointer viewPointer )
    {
	    if ( viewPointer != null )
	    {
		    // �������� ����� �������
		    viewPointer.X2 = 90;
		    viewPointer.Y2 = 140;
		    // �������� ��������������
		    viewPointer.AutoSorted = true;

		    // �������� ��������� ������ ����������� ����������� �������
		    IText adText = viewPointer.AdditionalText;
    		
		    if ( adText != null )
			    // �������� �����
			    SetTextSmallFont( ref adText, "(1)", 7 );
    		
		    // ��������� ���������
		    viewPointer.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // �������� � �������������� ������� ����������� �������
    // ---
    void ViewPointerWork()
    {
	    // �������� ��������� �������� �����������
	    ISymbols2DContainer symbCont = GetSymbols2DContainer();
    	
	    if ( symbCont != null )
	    {
		    // �������� ��������� ������� ����������� �������
		    IViewPointers viewPointsCol = symbCont.ViewPointers;
    		
		    if ( viewPointsCol != null )
		    {
			    // �������� ������� ����������� �������
			    IViewPointer newPointer = viewPointsCol.Add();
    			
			    if ( newPointer != null && doc2D != null )
			    {
				    // ������� ������� ����������� �������
				    bool create = CreateViewPointer( ref newPointer );
    				
				    // �������� �������� �������
				    int objRef = newPointer.Reference;
				    // ���������� ��������� ������
            doc2D.ksLightObj( objRef, 1/*�������� ���������*/ );
    				
            if ( create && kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
				    {
					    // �������� ������ �� ��������� �� ���������
					    IViewPointer viewPointer = viewPointsCol.get_ViewPointer( objRef );				
					    // ������������� ������� ����������� �������
					    EditViewPointer( ref viewPointer );
				    }
            doc2D.ksLightObj( objRef, 0/*��������� ���������*/ );
			    }
		    }
	    }
    }
    #endregion


    #region ������ �����
    //-------------------------------------------------------------------------------
    // ������ ����� � ������� ������� �����
    // ---
    void SetToleranceText( ref IToleranceParam tolPar )
    {
	    if ( tolPar != null )
	    {
		    // �������� ��������� ������� � ������� ������� �����
		    ITable tolTable = tolPar.Table;
    		
		    if ( tolTable != null )
		    {
			    // �������� 3 ������� (1 ��� ����)
			    tolTable.AddColumn( -1, true/*������*/ );
			    tolTable.AddColumn( -1, true/*������*/ );
			    tolTable.AddColumn( -1, true/*������*/ );
    			
			    // �������� ����� � 1-� ������
			    ITableCell cell = tolTable.get_Cell( 0, 0 );
			    ITextLine txt = null;
    			
			    if ( cell != null )
			    {
				    txt = (ITextLine)cell.Text;
    				
				    if ( txt != null )
					    txt.Str = "@22~";
			    }
    			
			    // �������� ����� �� 2-� ������
			    cell = tolTable.get_Cell( 0, 1 );
    			
			    if ( cell != null )
			    {
				    txt = (ITextLine)cell.Text;
    				
				    if ( txt != null )
					    txt.Str = "@2~";
			    }

			    // �������� ����� � 3-� ������
			    cell = tolTable.get_Cell( 0, 2 );
    			
			    if ( cell != null )
			    {
				    txt = (ITextLine)cell.Text;
    				
				    if ( txt != null )
					    txt.Str = "B";
			    }

			    // �������� ����� � 4-� ������
			    cell = tolTable.get_Cell( 0, 3 );
    			
			    if ( cell != null )
			    {
				    txt = (ITextLine)cell.Text;
    				
				    if ( txt != null )
					    txt.Str = "@30~";
			    }
		    }
	    }
    }


    //-------------------------------------------------------------------------------
    // �������� ������� �����
    // ---
    bool CreateTolerance( ref ITolerance tolerance )
    {
      bool res = false;

	    if ( tolerance != null )
	    {
		    // �������� ��������� �����������
		    IBranchs branchs = (IBranchs)tolerance;

		    if ( branchs != null )
		    {
			    // ������ ����� ��������
			    branchs.X0 = 100;
			    branchs.Y0 = 150;
			    // �������� 2 �����������
			    branchs.AddBranchByPoint( -1, 100, 120 );
			    branchs.AddBranchByPoint( -1, 50, 155 );
		    }

		    // �������� ��������� ���������� ������� �����
		    IToleranceParam tolPar = (IToleranceParam)tolerance;
    		
		    if ( tolPar != null )
		    {
			    // ������� ����� � �������
			    SetToleranceText( ref tolPar );
			    // ��������� ������� ����� ������������ ������� - ����� ����������
			    tolPar.BasePointPos = Kompas6Constants.ksTablePointEnum.ksTPBottomCenter;
		    }
		    // ��� ������� 1-�� ����������� - �����������
		    tolerance.set_ArrowType( 0, false );
		    // ��������� 1-�� ����������� ������������ ������� - ����� ����������
		    tolerance.set_BranchPos( 0, Kompas6Constants.ksTablePointEnum.ksTPBottomCenter );
		    // ��� ������� 2-�� ����������� - �������
		    tolerance.set_ArrowType( 1, true );
		    // ��������� 2-�� ����������� ������������ ������� - ����� ����������
		    tolerance.set_BranchPos( 1, Kompas6Constants.ksTablePointEnum.ksTPLeftCenter );
		    // ��������� ���������
		    res = tolerance.Update();
	    }
      return res;
    }


    //-------------------------------------------------------------------------------
    // �������������� ������� �����
    // ---
    void EditTolerance( ref ITolerance tolerance )
    {	
	    if ( tolerance != null )
	    {
		    // ����������������� ���������� ������� �����
		    IToleranceParam tolPar = (IToleranceParam)tolerance;

		    if ( tolPar != null )
		    {
			    // �������� ��������� ������� � ������� ������� �����
			    ITable tolTable = tolPar.Table;
    			
			    if ( tolTable != null )
			    {
				    // �������� ����� �� 2-� ������
				    ITableCell cell = tolTable.get_Cell( 0, 1 );
    				
				    if ( cell != null )
				    {
					    ITextLine txt = (ITextLine)cell.Text;
    					
					    if ( txt != null )
						    txt.Str = "@2~15";
				    }
			    }
			    // ������ ������� ��������������
			    tolPar.Vertical = true;
		    }

		    // �������� ��������� �����������
		    IBranchs branchs = (IBranchs)tolerance;

		    if ( branchs != null )
		    {
			    // ������� �����������
			    branchs.DeleteBranch( 0 );
			    // �������� ����� �����������
			    branchs.AddBranchByPoint( -1, 130, 120 );
		    }
		    tolerance.set_ArrowType( 1, false );
		    tolerance.set_BranchPos( 1, Kompas6Constants.ksTablePointEnum.ksTPBottomCenter );
		    // ��������� ���������
		    tolerance.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // �������� � �������������� ������� �����
    // ---
    void ToleranceWork()
    {
	    // �������� ��������� �������� �����������
	    ISymbols2DContainer symbCont = GetSymbols2DContainer();
    	
	    if ( symbCont != null )
	    {
		    // �������� ��������� �������� �����
		    ITolerances tolerancesCol = symbCont.Tolerances;
    		
		    if ( tolerancesCol != null )
		    {
			    // �������� ������ �����
			    ITolerance newTol = tolerancesCol.Add();
    			
			    if ( newTol != null )
			    {
				    // ������� ������ �����
				    bool create = CreateTolerance( ref  newTol );
    				
				    // �������� �������� �������
				    int objRef = newTol.Reference;
				    // ���������� ��������� ������
            doc2D.ksLightObj( objRef, 1/*�������� ���������*/ );
    				
            if ( create && kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
				    {
					    // �������� ������ �� ��������� �� ���������
					    ITolerance tolerance = tolerancesCol.get_Tolerance( objRef );				
					    // ������������� ������ �����
					    EditTolerance( ref tolerance );
				    }
            doc2D.ksLightObj( objRef, 0/*��������� ���������*/ );
			    }
		    }
	    }
    }
    #endregion


    #region ��������� �� �������� ����
    //-------------------------------------------------------------------------------
    // �������� ��������� ����������� ��������
    // ---
    IDrawingContainer GetDrawingContainer()
    {
	    if ( doc != null )
	    {
		    // �������� �������� ����� � �����
		    IViewsAndLayersManager mng = doc.ViewsAndLayersManager;

		    if ( mng != null )
		    {
			    // �������� ��������� �����
			    IViews viewsCol = mng.Views;

			    if ( viewsCol != null )
			    {
				    // �������� �������� ���
				    IView view = viewsCol.ActiveView;

				    if ( view != null )
				    {
					    // �������� ��������� ����������� ��������
					    IDrawingContainer drawCont = (IDrawingContainer)view;
					    return drawCont;
				    }
			    }
		    }
	    }
	    return null;
    }


    //-------------------------------------------------------------------------------
    // ������ ��������� � ����������� ������� ����� �������
    // ---
    void GetLeaderPar( ref ILeader leader )
    {
	    if ( leader != null )
	    {
		    double x0 = 0, y0 = 0;

		    // �������� ���������� ����� ��������
		    IBranchs branchs = (IBranchs)leader;

		    if ( branchs != null )
		    {
			    x0 = branchs.X0;
			    y0 = branchs.Y0;
		    }

		    // ������������ ��������� ��� ������
		    // "������� ����� �������"
		    string buf = LoadString( "IDS_LEADER" );
        buf += "\n\r" + LoadString( "IDS_POINT" ) + "\n\r";
        buf += string.Format( LoadString("IDS_COORDS"), x0, y0 );
		    // "\n������� ��������� �� �������: "
        buf += "\n\r" + LoadString( "IDS_ARROUND" );

		    if ( leader.Arround )
			    // "�������"
          buf += LoadString( "IDS_ON" );
		    else
			    // "��������"
          buf += LoadString( "IDS_OFF" );

		    // "\n������ �����������: "
        buf += "\n\r" + LoadString( "IDS_BRANCHBEGIN" );

		    if ( leader.get_BranchBegin(0) )
			    // "�� ������ �����"
          buf += LoadString( "IDS_BEGINSHELF" );
		    else
			    // "�� ����� �����"
          buf = buf + LoadString( "IDS_ENDSHELF" );

		    // "\n������� �������������� �����������: "
        buf += "\n\r" + LoadString( "IDS_PARALLEL" );

		    if ( leader.ParallelBranch )
			    // "�������"
          buf += LoadString( "IDS_ON" );
		    else
			    // "��������"
          buf += LoadString( "IDS_OFF" );
    		
		    // ������ ��������� 
		    kompas.ksMessage( buf );
	    }
    }


    //-------------------------------------------------------------------------------
    // ������ ��������� � ����������� ����� ����������
    // ---
    void GetMarkLeaderPar( ref IMarkLeader markLeader )
    {
	    if ( markLeader != null )
	    {
		    double x0 = 0, y0 = 0;
		    long count = 0;
    	
		    IBranchs branchs = (IBranchs)markLeader;
    		
		    if ( branchs != null )
		    {
			    // �������� ���������� ����� ��������
			    x0 = branchs.X0;
			    y0 = branchs.Y0;
			    // �������� ���������� �����������
			    count = branchs.BranchCount;
		    }

		    // ������������ ��������� ��� ������
		    // "���� ����������"
        string buf = LoadString( "IDS_MARKLEADER" );
		    // "\n����� ��������:\nX0 = %4.2f, Y0 = %4.2f"
        buf += "\n\r" + LoadString( "IDS_POINT" ) + "\n\r";
        buf += string.Format( LoadString("IDS_COORDS"), x0, y0 );
		    // "\n���������� �����������: %d"
        buf += "\n\r" + string.Format( LoadString("IDS_BRANCHCOUNT"), count );

		    // ������ ��������� 
        kompas.ksMessage( buf );
	    }
    }


    //-------------------------------------------------------------------------------
    // ������ ��������� � ����������� ����� ���������
    // ---
    void GetBrandLeaderPar( ref IBrandLeader brandLeader )
    {
	    if ( brandLeader != null )
	    {
		    // ������������ ���������
        string buf = LoadString( "IDS_BRANDLEADER" );
		    // ������ ��������� 
		    kompas.ksMessage( buf );
	    }
    }


    //-------------------------------------------------------------------------------
    // ������ ��������� � ����������� ����� ���������
    // ---
    void GetChangeLeaderPar( ref IChangeLeader changeLeader )
    {
	    if ( changeLeader != null )
	    {
		    // ������������ ���������
        string buf = LoadString( "IDS_CHANLEADER" );
		    // ������ ��������� 
		    kompas.ksMessage( buf );
	    }
    }


    //-------------------------------------------------------------------------------
    // ������ ��������� � ����������� ����������� �������������
    // ---
    void GetRoughPar( ref IRough rough )
    {
	    if ( rough != null )
	    {
		    // ������������ ���������
        string buf = LoadString( "IDS_ROUGH" );
		    // ������ ��������� 
		    kompas.ksMessage( buf );
	    }
    }


    //-------------------------------------------------------------------------------
    // ������ ��������� � ����������� ����������� ����
    // ---
    void GetBasePar( ref IBase pBase )
    {
	    if ( pBase != null )
	    {
		    // ������������ ���������
        string buf = LoadString( "IDS_BASE" );
		    // ������ ��������� 
		    kompas.ksMessage( buf );
	    }
    }


    //-------------------------------------------------------------------------------
    // ������ ��������� � ����������� ����� �������/�������
    // ---
    void GetCutLinePar( ref ICutLine cutLine )
    {
	    if ( cutLine != null )
	    {
		    // ������������ ���������
        string buf = LoadString( "IDS_CUTLINE" );
		    // ������ ��������� 
		    kompas.ksMessage( buf );
	    }
    }


    //-------------------------------------------------------------------------------
    // ������ ��������� � ����������� ������� ����������� �������
    // ---
    void GetViewPointerPar( ref IViewPointer viewPointer )
    {
	    if ( viewPointer != null )
	    {
		    // ������������ ���������
        string buf = LoadString( "IDS_VIEWPOINTER" );
		    // ������ ��������� 
		    kompas.ksMessage( buf );
	    }
    }


    //-------------------------------------------------------------------------------
    // ������ ��������� � ����������� ������� �����
    // ---
    void GetTolerancePar( ref ITolerance tolerance )
    {
	    if ( tolerance != null )
	    {
		    // ������������ ���������
        string buf = LoadString( "IDS_TOLERANCE" );
		    // ������ ��������� 
		    kompas.ksMessage( buf );
	    }
    }


    //-------------------------------------------------------------------------------
    // ��������� �� ������� �������� ����
    // ---
    void ObjectsNavigation()
    {
	    // �������� ��������� ����������� ��������
	    IDrawingContainer drawCont = GetDrawingContainer();

	    if ( drawCont != null )
	    {
		    // �������� ������ �������� 
        Array arr = (Array)drawCont.get_Objects( 0/*��� �������*/ );

        // ���� ������ ����
		    if ( arr != null )
		    {    	
			    for ( long j = 0; j < arr.Length; j++ )
			    {
				    // �������� ������� �� �������
				    stdole.IDispatch pObj = (stdole.IDispatch)arr.GetValue( j );
    				
				    if ( pObj != null )
				    {
					    // �������� ��������� ������������ �������
					    IDrawingObject pDrawObj = (IDrawingObject)pObj;
    					
					    if ( pDrawObj != null )
					    {
						    // �������� ��� �������
						    long type = (int)pDrawObj.DrawingObjectType;
						    // �������� �������� �������
						    int objRef = pDrawObj.Reference;
						    // ���������� ������
						    doc2D.ksLightObj( objRef, 1/*��������*/ );

						    // � ����������� �� ���� ������� ��������� ��� ������� ���� ��������
						    switch( type )
						    {
							    // ������� ����� �������
							    case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrLeader:
							    {
								    ILeader leader = (ILeader)pDrawObj;
								    GetLeaderPar( ref leader );
								    break;
							    }

							    // ���� ����������
							    case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrMarkerLeader:
							    {
								    IMarkLeader markLeader = (IMarkLeader)pDrawObj;
								    GetMarkLeaderPar( ref markLeader );
								    break;
							    }

							    // ���� ���������
							    case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrBrandLeader:
							    {
								    IBrandLeader brandLeader = (IBrandLeader)pDrawObj;
								    GetBrandLeaderPar( ref brandLeader );
								    break;
							    }

							    // ���� ���������
							    case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrChangeLeader:
							    {
								    IChangeLeader changeLeader = (IChangeLeader)pDrawObj;
								    GetChangeLeaderPar( ref changeLeader );
								    break;
							    }

							    // ����������� �������������
							    case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrRough:
							    {
								    IRough rough = (IRough)pDrawObj;
								    GetRoughPar( ref rough );
								    break;
							    }

							    // ����������� ����
							    case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrBase:
							    {
								    IBase pBase = (IBase)pDrawObj;
								    GetBasePar( ref pBase );
								    break;
							    }

							    // ����� �������/�������
							    case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrCut:
							    {
								    ICutLine cutLine = (ICutLine)pDrawObj;
								    GetCutLinePar( ref cutLine );
								    break;
							    }

							    // ������� ����������� �������
							    case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrWPointer:
							    {
								    IViewPointer viewPointer = (IViewPointer)pDrawObj;
								    GetViewPointerPar( ref viewPointer );
								    break;
							    }

							    // ������ �����
							    case (int)Kompas6Constants.DrawingObjectTypeEnum.ksDrTolerance:
							    {
								    ITolerance tolerance = (ITolerance)pDrawObj;
								    GetTolerancePar( ref tolerance );
								    break;
							    }
						    }
						    // ������ ���������
						    doc2D.ksLightObj( objRef, 0/*���������*/ );
					    } 
				    }
			    }
		    }
	    }
    }
    #endregion


    #region COM Registration
    // ��� ������� ����������� ��� ����������� ������ ��� COM
    // ��� ��������� � ����� ������� ���������� ������ Kompas_Library,
    // ������� ������������� � ���, ��� ����� �������� ����������� ������,
    // � ����� �������� ��� InprocServer32 �� ������, � ��������� ����.
    // ��� ��� �������� ��� ����, ����� ����� ����������� ����������
    // ���������� �� ������� ActiveX.
    [ComRegisterFunction]
    public static void RegisterKompasLib(Type t)
    {
      try
      {
        RegistryKey regKey = Registry.LocalMachine;
        string keyName = @"SOFTWARE\Classes\CLSID\{" + t.GUID.ToString() + "}";
        regKey = regKey.OpenSubKey(keyName, true);
        regKey.CreateSubKey("Kompas_Library");
        regKey = regKey.OpenSubKey("InprocServer32", true);
        regKey.SetValue(null, System.Environment.GetFolderPath(Environment.SpecialFolder.System) + @"\mscoree.dll");
        regKey.Close();
      }
      catch (Exception ex)
      {
        MessageBox.Show(string.Format("��� ����������� ������ ��� COM-Interop ��������� ������:\n{0}", ex));
      }
    }
		
    // ��� ������� ������� ������ Kompas_Library �� �������
    [ComUnregisterFunction]
    public static void UnregisterKompasLib(Type t)
    {
      RegistryKey regKey = Registry.LocalMachine;
      string keyName = @"SOFTWARE\Classes\CLSID\{" + t.GUID.ToString() + "}";
      RegistryKey subKey = regKey.OpenSubKey(keyName, true);
      subKey.DeleteSubKey("Kompas_Library");
      subKey.Close();
    }
    #endregion 
	}
}
