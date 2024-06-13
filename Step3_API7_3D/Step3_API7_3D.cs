using Kompas6API5;
using KompasAPI7;
using Kompas6Constants;
using Kompas6Constants3D;
using System;
using System.Resources;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Windows.Forms;


namespace Step3_API7_3D
{
	public class Symbols3D
	{
    private KompasObject      kompas;
    private IApplication      appl;         // ��������� ����������
    private IKompasDocument3D doc;          // ��������� ��������� 3D � API7 
    private ksDocument3D      doc3D;        // ��������� ��������� 3D � API5
    private ResourceManager   resMng = new ResourceManager( typeof(Symbols3D) );   // �������� ��������
    private int               oType = (int)Kompas6Constants3D.Obj3dType.o3d_face;


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
          result = LoadString( "IDS_MENU1" );
          command = 1;
        break;

        case 2:
          result = LoadString( "IDS_MENU2" );
          command = 2;
        break;

        case 3:
          result = LoadString( "IDS_MENU3" );
          command = 3;
        break;

        case 4:
          result = LoadString( "IDS_MENU4" );
          command = 4;
        break;

        case 5:
          result = LoadString( "IDS_MENU5" );
          command = 5;
        break;

        case 6:
          result = LoadString( "IDS_MENU6" );
          command = 6;
        break;

        case 7:
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

      // �������� ��������� ��������� ��������� 3D � API7
      doc = (IKompasDocument3D)appl.ActiveDocument;

      if ( doc == null )
        return;

      // �������� ��������� ��������� ��������� 3D � API5
      doc3D = (ksDocument3D)kompas.ActiveDocument3D();

      if ( doc3D == null )
        return;

      switch ( command )
      {
        case 1: Rough3DWork();        break;  // �������� � �������������� ������������� 3D
        case 2: Base3DWork();         break;  // �������� � �������������� ����������� ���� 3D
        case 3: Leader3DWork();       break;  // �������� � �������������� �����-������� 3D
        case 4: BrandLeader3DWork();  break;  // �������� � �������������� ����� ��������� 3D
        case 5: MarkLeader3DWork();   break;  // �������� � �������������� ����� ���������� 3D
        case 6: Tolerance3DWork();    break;  // �������� � �������������� ������� ����� 3D
      }
    }


    #region ��������������� �������
    //-------------------------------------------------------------------------------
    // �������� ��������� ����������� 3D
    // ---
    ISymbols3DContainer GetSymbols3DContainer()
    {
      if ( doc != null )
      {
        // �������� ��������� ����������� 3D
        return (ISymbols3DContainer)doc.TopPart;
      }
      return null;
    }


    //-----------------------------------------------------------------------------
    // ������� ����������
    // ---
    public bool UserFilterProc( object e )
    {
      ksEntity entity = (ksEntity)e;

      if( e != null && (oType == 0 || entity.type == oType) )
      {
        return true;
      }
      else
        return false;
    }


    //-------------------------------------------------------------------------------
    // ������ ��������� ����������� �� ������� �������� �������
    // ---
    void SetPosition( ref IRough3D rough, ref IModelObject face )
    {
	    if ( rough != null && face != null )
	    {
		    // ������������� ��������� ������� ������ �� API7 � API5
        ksEntity faceEnt = (ksEntity)kompas.TransferInterface( face, 
                           (int)Kompas6Constants.ksAPITypeEnum.ksAPI5Auto,
                           (int)Kompas6Constants3D.ksObj3dTypeEnum.o3d_entity );

		    if ( faceEnt != null )
		    { 
			    // �������� ��������� ���������� �����
			    ksFaceDefinition faceDef = (ksFaceDefinition)faceEnt.GetDefinition();

			    if ( faceDef != null )
			    {
				    // �������� ��������� �����
				    ksEdgeCollection edgeCol = (ksEdgeCollection)faceDef.EdgeCollection();

				    if ( edgeCol != null )
				    {
					    // �������� ��������� ���������� �����
					    ksEdgeDefinition edgeDef = (ksEdgeDefinition)edgeCol.First();

					    if ( edgeDef != null )
					    {
						    // �������� ��������� ���������� �������
						    ksVertexDefinition vertex = (ksVertexDefinition)edgeDef.GetVertex( true/*��������� �������*/ );

						    if ( vertex != null )
						    {
							    // ������������� ��������� ������� �� API5 � API7
                  IModelObject mObj = (IModelObject)kompas.TransferInterface( vertex, 
                                      (int)Kompas6Constants.ksAPITypeEnum.ksAPI7Dual, 0 );

							    if ( mObj != null )
								    // ���������� ��������� ��������� ����������� �� ��������� �������
								    rough.PositionObject = mObj;
						    }
					    }
				    }
			    }
		    }
	    }
    }


    //-------------------------------------------------------------------------------
    // �������� ����� ��� �������
    // ---
    string GetNewName( ref IFeature7 featObj, string strID )
    {
	    string res = "";

	    if ( featObj != null )
	    {
		    // ������������ ����� ��� ������� � ������
		    string newName =	LoadString( strID );
		    string name = featObj.Name;
        int pos = name.IndexOf( ":", 0, name.Length );
    		
		    if ( pos > 0 )
			    newName += name.Substring( pos + 1, name.Length - pos - 1 );

        res = newName;
	    }
	    return res;
    }


    //-------------------------------------------------------------------------------
    // 
    // ---
    Int32 RGB( Int32 r, Int32 g, Int32 b )
    {
      return ( (r | (g << 8)) | (b << 16) );
    }


    //-------------------------------------------------------------------------------
    // ������ ���������� ������
    // ---
    void AddTextItem( ref ITextLine line, string str, 
                      Kompas6Constants.ksTextItemEnum type )
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
    #endregion


    #region ������������� 3D
    //-------------------------------------------------------------------------------
    // ���������� ��������� �������������
    // ---
    void SetRoughPars( ref IRough3D rough )
    {
	    if ( rough != null )
	    {
		    // �������� ��������� ���������� �������������
		    IRoughParams roughPars = (IRoughParams)rough;
    		
		    if ( roughPars != null )
		    {
			    // ������� ��������� �� �������
			    roughPars.ProcessingByContour = true;
			    // ����� ����� �������
			    roughPars.LeaderLength = 20;
			    // ���� ������� ����� �������
			    roughPars.LeaderAngle = 45;
    			
			    // �������� ��������� ������ ���������� �������������
			    IText txt1 = roughPars.RoughParamText;
    			
			    if ( txt1 != null )
				    txt1.Str = "1";
    			
			    // �������� ��������� ������ ������� ��������� �����������
			    IText txt2 = roughPars.ProcessText;
    			
			    if ( txt2 != null )
				    txt2.Str = "2";
    			
			    // �������� ��������� ������ ������� �����
			    IText txt3 = roughPars.BaseLengthText;
    			
			    if ( txt3 != null )
				    txt3.Str = "3";
		    }
	    }
    }


    //-------------------------------------------------------------------------------
    // ������� ������������� 3D
    // ---
    bool CreateRough3D( ref IRough3D rough )
    {
	    bool res = false;

	    if ( rough != null && doc3D != null )
	    {
			  // ������� � ��������� ������� ������ ��� ���������� �������������
        ksEntity baseObj = (ksEntity)doc3D.UserSelectEntity( null, "UserFilterProc", 
                                           LoadString("IDS_OBJ"), 0, this );

			  if ( baseObj != null )
			  {
          double x = 0, y = 0, z = 0;
				  // ������� ����� ��������
          doc3D.UserGetCursor( LoadString("IDS_POINT"), out x, out y, out z );

				  // ������������� ��������� ������� �� API5 � API7
          IModelObject mObj = (IModelObject)kompas.TransferInterface( baseObj, 
                              (int)Kompas6Constants.ksAPITypeEnum.ksAPI7Dual, 0 );

				  if ( mObj != null )
					  // ���������� �������� �������������
					  rough.SetBasePosition( x, y, z, mObj );

				  // ������ ������� ���������
				  rough.BasePlane = Kompas6Constants3D.ksObj3dTypeEnum.o3d_planeXOZ;

				  // ���������� ��������� �������������
				  SetRoughPars( ref rough );

				  // ��������� ���������
				  res = rough.Update();
			  }
		  }
	    return res;
    }
    

    //-------------------------------------------------------------------------------
    // ������������� ������������� 3D
    // ---
    void EditRough3D( ref IRough3D rough )
    {
	    if ( rough != null )
	    {
		    // �������� ��������� ���������� �������������
		    IRoughParams roughPars = (IRoughParams)rough;

		    if ( roughPars != null )
			    // ������ ������� ��������� �� �������
			    roughPars.ProcessingByContour = false;

		    // �������� ��������� ������� ������
		    IFeature7 featObj = (IFeature7)rough;

		    if ( featObj != null )
		    {	
			    // �������� ��� ������� � ������	
			    rough.Name = GetNewName( ref featObj, "IDS_ROUGH" ); 
			    // �������� ������� ������
			    IModelObject baseObj = rough.BaseObject;
			    // ������ ��������� ����������� �� ������� �������� �������
			    SetPosition( ref rough, ref baseObj );
		    }

		    // �������� ��������� ������� ����� �������
		    IColorParam7 colorPars = (IColorParam7)rough;

		    // �������� ���� �� �����
		    if ( colorPars != null )
			    colorPars.Color = RGB( 0, 0, 255 );

		    // ��������� ���������
		    rough.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // �������� � �������������� ������������� 3D
    // ---
    void Rough3DWork()
    {
	    // �������� ��������� ����������� 3D
	    ISymbols3DContainer symbCont = GetSymbols3DContainer();
    	
	    if ( symbCont != null )
	    {
		    // �������� ��������� �������������� 3D
		    IRoughs3D roughsCol = symbCont.Roughs3D;
    		
		    if ( roughsCol != null )
		    {
			    // �������� ����� ����������� ������������� 3D
			    IRough3D newRough = roughsCol.Add();
    			
			    if ( newRough != null )
			    {
				    // ������� ����������� ������������� 3D
				    if ( CreateRough3D(ref newRough) )
				    {	
					    string name = "";
					    // �������� ��������� ������� ������ ����������
					    IFeature7 feature = (IFeature7)newRough;
    					
					    // �������� ��� �������
					    if ( feature != null )
						    name = feature.Name;
    					
					    if ( kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
					    {
						    // �������� ������ �� ��������� �� �����
						    IRough3D rough = roughsCol.get_Rough3D( name );
						    // ������������� �������������
						    EditRough3D( ref rough );
					    }
				    }
				    else
					    // ������ ��������� "������ �� ������"
              kompas.ksMessage( LoadString("IDS_NOCREATE") );
			    }
		    }
	    }
    }
    #endregion


    #region ����������� ���� 3D
    //-------------------------------------------------------------------------------
    // ������� ����������� ����
    // ---
    bool CreateBase3D( ref IBase3D pBase ) 
    {
	    bool res = false;

	    if ( pBase != null && doc3D != null )
	    {
			  // ������� � ��������� ������� ������ ��� ���������� ����
        ksEntity baseObj = (ksEntity)doc3D.UserSelectEntity( null, "UserFilterProc", 
                                           LoadString("IDS_OBJ"), 0, this );
    		
			  if ( baseObj != null )
			  {
				  double x = 0, y = 0, z = 0;
				  // ������� ����� ��������
				  doc3D.UserGetCursor( LoadString("IDS_POINT"), out x, out y, out z );
    			
				  // ������������� ��������� ������� �� API5 � API7
          IModelObject mObj = (IModelObject)kompas.TransferInterface( baseObj, 
                              (int)Kompas6Constants.ksAPITypeEnum.ksAPI7Dual, 0 );
    			
				  if ( mObj != null )
					  // ���������� �������� ����
					  pBase.SetBranchBeginPoint( x, y, z, mObj );
    			
				  // ������ ������� ���������
				  pBase.BasePlane = Kompas6Constants3D.ksObj3dTypeEnum.o3d_planeXOZ;
    			
				  // ��������� ���������
				  res = pBase.Update();
			  }
		  }
	    return res;
    }


    //-------------------------------------------------------------------------------
    // �������������� ����������� ����
    // ---
    void EditBase3D( ref IBase3D pBase )
    {
	    if ( pBase != null )
	    {
		    // �������� ��������������
		    pBase.AutoSorted = false;

		    // �������� ��������� ������ ����������� ����
		    IText txt = pBase.Text;
    		
		    if ( txt != null )
		    {
			    txt.Str = "";
    			
			    ITextLine line = txt.Add();
    			
			    if ( line != null )
			    {
				    // �������� ������
				    AddTextItem( ref line, "A", Kompas6Constants.ksTextItemEnum.ksTItString );
				    // �������� ��������� ���� ����� (������ ������ ����� ������������
				    // ������ � ����� ���������
				    AddTextItem( ref line, "", Kompas6Constants.ksTextItemEnum.ksTItSBase );
				    // �������� ������ ������
				    AddTextItem( ref line, "1", Kompas6Constants.ksTextItemEnum.ksTItSLowerIndex );
				    // ��������� ��������� ���� �����
				    AddTextItem( ref line, "", Kompas6Constants.ksTextItemEnum.ksTItSEnd );
			    }
		    }

		    // ��� - ������������ ���������
		    pBase.DrawType = false;
		    double x = 0, y = 0, z = 0;
		    // �������� �������� ����� ����������� ����
		    pBase.GetBranchEndPoint( out x, out y, out z );
		    pBase.SetBranchEndPoint( x + 10, y + 10, z );

		    // �������� ��������� ������� ������
		    IFeature7 featObj = (IFeature7)pBase;

		    if ( featObj != null )
			    // �������� ��� ������� � ������
			    pBase.Name = GetNewName( ref featObj, "IDS_BASE" );

		    // �������� ��������� ������� ����� �������
		    IColorParam7 colorPars = (IColorParam7)pBase;
    		
		    // �������� ���� �� ������
		    if ( colorPars != null )
			    colorPars.Color = RGB( 0, 0, 0 );


		    pBase.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // �������� � �������������� ����������� ����
    // ---
    void Base3DWork()
    {
	    // �������� ��������� ����������� 3D
	    ISymbols3DContainer symbCont = GetSymbols3DContainer();
    	
	    if ( symbCont != null )
	    {
		    // �������� ��������� ����������� ���� 3D
		    IBases3D basesCol = symbCont.Bases3D;
    		
		    if ( basesCol != null )
		    {
			    // �������� ����� ����������� ���� 3D
			    IBase3D newBase = basesCol.Add();
    			
			    if ( newBase != null )
			    {
				    // ������� ����������� ���� 3D
				    if ( CreateBase3D(ref newBase) )
				    {	
					    string name = "";
					    // �������� ��������� ������� ������ ����������
					    IFeature7 feature = (IFeature7)newBase;
    					
					    // �������� ��� �������
					    if ( feature != null )
						    name = feature.Name;
    					
					    if ( kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
					    {
						    // �������� ������ �� ��������� �� �����
						    IBase3D pBase = basesCol.get_Base3D( name );
						    // ������������� ����������� ����
						    EditBase3D( ref pBase );
					    }
				    }
				    else
					    // ������ ��������� "������ �� ������"
              kompas.ksMessage( LoadString("IDS_NOCREATE") );
			    }
		    }
	    }
    }
    #endregion


    #region �����-������� 3D
    //-------------------------------------------------------------------------------
    // ���������� ������ �����-�������
    // ---
    void SetLeader3DTexts( ref IBaseLeader3D baseLeader )
    {
	    ILeader leader = (ILeader)baseLeader;

	    if ( leader != null )
	    {
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
	    }
    }


    //-------------------------------------------------------------------------------
    // ������� ����������� �����-�������
    // ---
    void CreateBranchsLeader( ref IBaseLeader3D leader )
    {
	    if ( leader != null && doc3D != null )
	    {
			  // ������� � ��������� ������� ������
        ksEntity baseObj = (ksEntity)doc3D.UserSelectEntity( null, "UserFilterProc", 
                                           LoadString("IDS_OBJ"), 0, this );
    		
			  double x =0 , y = 0, z = 0, x1 = 0, y1 = 0, z1 = 0;
			  // ������� ����� ������ �����.
			  // ����� ������ ������ �� ������� �������
			  doc3D.UserGetCursor( LoadString("IDS_BEGINBRANCH"), out x, out y, out z );
    		
			  // ������� ������, �������� ��������� �����������
        ksEntity posObj = (ksEntity)doc3D.UserSelectEntity( null, "UserFilterProc", 
                                          LoadString("IDS_POSOBJ"), 0, this );
    		
			  if ( baseObj != null && posObj != null )
			  {
				  // ������� �����, �������� �������� ��������
				  // ����� ������ ������ � ��������� �����������
				  doc3D.UserGetCursor( LoadString("IDS_BEGINSHELF"), out x1, out y1, out z1 );
				  // ������������� ��������� ������� �� API5 � API7
          IModelObject mObj1 = (IModelObject)kompas.TransferInterface( baseObj, 
                               (int)Kompas6Constants.ksAPITypeEnum.ksAPI7Dual, 0 );
          IModelObject mObj2 = (IModelObject)kompas.TransferInterface( posObj, 
                               (int)Kompas6Constants.ksAPITypeEnum.ksAPI7Dual, 0 );
    			
				  if ( mObj1 != null && mObj2 != null )
				  {	
					  // �������� ��������� �����������
					  IBranchs3D branchs3D = (IBranchs3D)leader;
    				
					  // �������� �����������
					  if ( branchs3D != null )
						  branchs3D.AddBranchByPoint( x, y, z, mObj1 );
    				
					  // ���������� ������, �������� ��������� �����������
					  leader.PositionObject = mObj2;
					  leader.SetPosition( x1, y1, z1 );
				  }
			  }
		  }
    }


    //-------------------------------------------------------------------------------
    // ������� �����-�������
    // ---
    bool CreateLeader3D( ref IBaseLeader3D leader ) 
    {
	    bool res = false;

	    if ( leader != null )
	    {
		    // ������� ����������� �����-�������
		    CreateBranchsLeader( ref leader );
		    // ���������� ������ �����-�������
		    SetLeader3DTexts( ref leader );
		    // ��������� ���������
		    res = leader.Update();
	    }
	    return res;
    }


    //-------------------------------------------------------------------------------
    // �������� ��������� �����-�������
    // ---
    void EditLeaderPars( ref IBaseLeader3D baseLeader )
    {
	    if ( baseLeader != null )
	    {
		    ILeader leader = (ILeader)baseLeader;

		    if ( leader != null )
		    {
			    // ��� ������ - ���� ����������
			    leader.SignType = Kompas6Constants.ksLeaderSignEnum.ksLGlueSign;
			    // �������� ������� ���������� �� �������
			    leader.Arround = true;
    			
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
		    }
	    }
    }


    //-------------------------------------------------------------------------------
    // �������� ������������� �����������
    // ---
    void LeaderAddBranch( ref IBaseLeader3D leader ) 
    {
	    if ( leader != null && doc3D != null )
	    {
			  // ������� � ��������� ������� ������
        ksEntity baseObj = (ksEntity)doc3D.UserSelectEntity( null, "UserFilterProc", 
                                           LoadString("IDS_OBJ"), 0, this );
    		
			  double x = 0, y = 0, z = 0;
			  // ������� ����� ������ �����
			  doc3D.UserGetCursor( LoadString("IDS_BEGINBRANCH"), out x, out y, out z );
    		
			  if ( baseObj != null )
			  {
				  // ������������� ��������� ������� �� API5 � API7
          IModelObject mObj = (IModelObject)kompas.TransferInterface( baseObj, 
                              (int)Kompas6Constants.ksAPITypeEnum.ksAPI7Dual, 0 );
    			
				  if ( mObj != null )
				  {	
					  // �������� ��������� �����������
					  IBranchs3D branchs3D = (IBranchs3D)leader;
    				
					  // �������� �����������
					  if ( branchs3D != null )
						  branchs3D.AddBranchByPoint( x, y, z, mObj );
				  }
			  }
		  }
    }


    //-------------------------------------------------------------------------------
    // ������������� �����-�������
    // ---
    void EditLeader3D( ref IBaseLeader3D leader ) 
    {
	    if ( leader != null )
	    {
		    // �������� ������������� �����������
		    LeaderAddBranch( ref leader );
		    // �������� ��������� �����-�������
		    EditLeaderPars( ref leader );

		    // �������� ��������� ������� ������
		    IFeature7 featObj = (IFeature7)leader;
    		
		    if ( featObj != null )
			    // �������� ��� ������� � ������
			    leader.Name = GetNewName( ref featObj, "IDS_LEADER" );
    		
		    // �������� ��������� ������� ����� �������
		    IColorParam7 colorPars = (IColorParam7)leader;
    		
		    // �������� ���� �� ����������
		    if ( colorPars != null )
			    colorPars.Color = RGB( 123, 40, 0 );

		    // ��������� ���������
		    leader.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // �������� � �������������� �����-������� 3D
    // ---
    void Leader3DWork()
    {
	    // �������� ��������� ����������� 3D
	    ISymbols3DContainer symbCont = GetSymbols3DContainer();
    	
	    if ( symbCont != null )
	    {
		    // �������� ��������� �����-������� 3D
		    ILeaders3D leadsCol = symbCont.Leaders3D;
    		
		    if ( leadsCol != null )
		    {
			    // �������� ����� �����-������� 3D
			    IBaseLeader3D newLeader = leadsCol.Add( Kompas6Constants3D.ksObj3dTypeEnum.
                                                  o3d_leader3D );
    			
			    if ( newLeader != null )
			    {
				    // ������� �����-������� 3D
				    if ( CreateLeader3D( ref newLeader) )
				    {	
					    string name = "";
					    // �������� ��������� ������� ������ ����������
					    IFeature7 feature = (IFeature7)newLeader;
    					
					    // �������� ��� �������
					    if ( feature != null )
						    name = feature.Name;
    					
					    if ( kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
					    {
						    // �������� ������ �� ��������� �� �����
						    IBaseLeader3D leader = leadsCol.get_Leader3D( name );
						    // ������������� �����-�������
						    EditLeader3D( ref leader );
					    }
				    }
				    else
					    // ������ ��������� "������ �� ������"
              kompas.ksMessage( LoadString("IDS_NOCREATE") );
			    }
		    }
	    }
    }
    #endregion


    #region ���� ��������� 3D
    //-------------------------------------------------------------------------------
    // ������� ���� ��������� 3D
    // ---
    bool CreateBrandLeader3D( ref IBaseLeader3D baseLeader )
    {
	    bool res = false;
    	
	    if ( baseLeader != null )
	    {
		    // ������� �����������
		    CreateBranchsLeader( ref baseLeader );

		    // �������� ��������� ����� ���������
		    IBrandLeader brandLeader = (IBrandLeader)baseLeader;

		    if ( brandLeader != null )
			    // �����������
			    brandLeader.Direction = false;
    			
		    // ��������� ���������
		    res = baseLeader.Update();
	    }
	    return res;
    }


    //-------------------------------------------------------------------------------
    // ������������� ���� ���������
    // ---
    void EditBrandLeader3D( ref IBaseLeader3D leader ) 
    {
	    if ( leader != null )
	    {
		    // �������� ������������� �����������
		    LeaderAddBranch( ref leader );

		    // ���������� ��������� ����� ���������
		    IBrandLeader brandLeader = (IBrandLeader)leader;
    		
		    if ( brandLeader != null )
		    {
			    // �������� ��������� ������ �����������
			    IText des = brandLeader.Designation;
    			
			    if ( des != null )
				    // �������� �����
				    des.Str = LoadString( "IDS_MARK1" );
		    }

		    // �������� ��������� ������� ������
		    IFeature7 featObj = (IFeature7)leader;
    		
		    if ( featObj != null )
			    // �������� ��� ������� � ������
			    leader.Name = GetNewName( ref featObj, "IDS_BRAND" );
    		
		    // �������� ��������� ������� ����� �������
		    IColorParam7 colorPars = (IColorParam7)leader;
    		
		    // �������� ���� �� ����������
		    if ( colorPars != null )
			    colorPars.Color = RGB( 128, 0, 128 );

		    // ��������� ���������
		    leader.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // �������� � �������������� ����� ���������
    // ---
    void BrandLeader3DWork()
    {
	    // �������� ��������� ����������� 3D
	    ISymbols3DContainer symbCont = GetSymbols3DContainer();
    	
	    if ( symbCont != null )
	    {
		    // �������� ��������� �����-������� 3D
		    ILeaders3D leadsCol = symbCont.Leaders3D;
    		
		    if ( leadsCol != null )
		    {
			    // �������� ����� ���� ��������� 3D
			    IBaseLeader3D newBrand = leadsCol.Add( Kompas6Constants3D.ksObj3dTypeEnum.
                                                 o3d_brandLeader3D );
    			
			    if ( newBrand != null )
			    {
				    // ������� ���� ��������� 3D
				    if ( CreateBrandLeader3D(ref newBrand) )
				    {	
					    string name = "";
					    // �������� ��������� ������� ������ ����������
					    IFeature7 feature = (IFeature7)newBrand;
    					
					    // �������� ��� �������
					    if ( feature != null )
						    name = feature.Name;
    					
					    if ( kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
					    {
						    // �������� ������ �� ��������� �� �����
						    IBaseLeader3D brandLeader = leadsCol.get_Leader3D( name );
						    // ������������� ���� ���������
						    EditBrandLeader3D( ref brandLeader );
					    }
				    }
				    else
					    // ������ ��������� "������ �� ������"
              kompas.ksMessage( LoadString("IDS_NOCREATE") );
			    }
		    }
	    }
    }
    #endregion


    #region ���� ���������� 3D
    //-------------------------------------------------------------------------------
    // ������� ���� ����������
    // ---
    bool CreateMarkLeader3D( ref IBaseLeader3D baseLeader )
    {
	    bool res = false;

	    if ( baseLeader !=null )
	    {
		    // ������� �����������
		    CreateBranchsLeader( ref baseLeader );

		    // �������� ��������� ����� ����������
		    IMarkLeader markLeader = (IMarkLeader)baseLeader;

		    if ( markLeader != null )
		    {	
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
		    }
		    // ��������� ���������
		    res = baseLeader.Update();
	    }
	    return res;
    }


    //-------------------------------------------------------------------------------
    // ������������� ���� ����������
    // ---
    void EditMarkLeader3D( ref IBaseLeader3D baseLeader )
    {
	    if ( baseLeader != null )
	    {
		    // �������� ������������� �����������
		    LeaderAddBranch( ref baseLeader );

		    // �������� ��������� ����� ����������
		    IMarkLeader markLeader = (IMarkLeader)baseLeader;
    		
		    if ( markLeader != null )
		    {
			    // �������� ��������� ������ �����������
			    IText des = markLeader.Designation;
    			
			    if ( des != null )
				    // �������� �����
				    des.Str = "10";
		    }

		    // �������� ��������� ������� ������
		    IFeature7 featObj = (IFeature7)baseLeader;
    		
		    if ( featObj != null )
			    // �������� ��� ������� � ������
			    baseLeader.Name = GetNewName( ref featObj, "IDS_BRAND" );
    		
		    // �������� ��������� ������� ����� �������
		    IColorParam7 colorPars = (IColorParam7)baseLeader;
    		
		    // �������� ���� �� ���������
		    if ( colorPars != null )
			    colorPars.Color = RGB( 204, 153, 255 );
    		
		    // ��������� ���������
		    baseLeader.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // �������� � �������������� ����� ���������� 3D
    // ---
    void MarkLeader3DWork()
    {
	    // �������� ��������� ����������� 3D
	    ISymbols3DContainer symbCont = GetSymbols3DContainer();
    	
	    if ( symbCont != null )
	    {
		    // �������� ��������� �����-������� 3D
		    ILeaders3D leadsCol = symbCont.Leaders3D;
    		
		    if ( leadsCol != null )
		    {
			    // �������� ����� ���� ���������� 3D
			    IBaseLeader3D newMark = leadsCol.Add( Kompas6Constants3D.ksObj3dTypeEnum.
                                                o3d_markLeader3D );
    			
			    if ( newMark != null )
			    {
				    // ������� ���� ���������� 3D
				    if ( CreateMarkLeader3D(ref newMark) )
				    {	
					    string name = "";
					    // �������� ��������� ������� ������ ����������
					    IFeature7 feature = (IFeature7)newMark;
    					
					    // �������� ��� �������
					    if ( feature != null )
						    name = feature.Name;

					    if ( kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
					    {
						    // �������� ������ �� ��������� �� �����
						    IBaseLeader3D markLeader = leadsCol.get_Leader3D( name );
						    // ������������� ���� ����������
						    EditMarkLeader3D( ref markLeader );
					    }
				    }
				    else
					    // ������ ��������� "������ �� ������"
              kompas.ksMessage( LoadString("IDS_NOCREATE") );
			    }
		    }
	    }
    }
    #endregion


    #region ������ ����� 3D
    //-------------------------------------------------------------------------------
    // ������� ����������� ��� �������� �����
    // ---
    void CreateToleranceBranch( ref ITolerance3D tolerance )
    {
	    if ( tolerance != null && doc3D != null )
	    {
			  // ������� � ��������� ������, �� ������� ����� ��������� �����
        ksEntity baseObj = (ksEntity)doc3D.UserSelectEntity( null, "UserFilterProc", 
                                           LoadString("IDS_OBJ"), 0, this );
    		
			  double x = 0, y = 0, z = 0, x1 = 0, y1 = 0, z1 = 0;
			  // ������� ����� ����� �����
			  // ����� ������ ������ �� ������� �������
			  doc3D.UserGetCursor( LoadString("IDS_ENDPOINT"), out x, out y, out z );
    		
			  // ������� ������, �������� ��������� �����������
        ksEntity posObj = (ksEntity)doc3D.UserSelectEntity( null, "UserFilterProc", 
                                          LoadString("IDS_POSOBJ"), 0, this );
    		
			  if ( baseObj != null && posObj != null )
			  {
				  // ������� ����� ������� �������
				  doc3D.UserGetCursor( LoadString("IDS_TABLE"), out x1, out y1, out z1 );
				  // ������������� ��������� ������� �� API5 � API7
          IModelObject mObj1 = (IModelObject)kompas.TransferInterface( baseObj, 
                               (int)Kompas6Constants.ksAPITypeEnum.ksAPI7Dual, 0 );
          IModelObject mObj2 = (IModelObject)kompas.TransferInterface( posObj, 
                               (int)Kompas6Constants.ksAPITypeEnum.ksAPI7Dual, 0 );
    			
				  if ( mObj1 != null && mObj2 != null )
				  {	
					  // �������� ��������� �����������
					  IBranchs3D branchs3D = (IBranchs3D)tolerance;
    				
					  // �������� �����������
					  if ( branchs3D != null )
						  branchs3D.AddBranchByPoint( x, y, z, mObj1 );
    				
					  // ���������� ������, �������� ��������� �����������
					  tolerance.PositionObject = mObj2;
					  // ���������� �����, �������� ��������� �������
					  tolerance.SetPosition( x1, y1, z1 );
				  }
			  }
		  }
    }


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
    				
				    if ( txt != null)
					    txt.Str = "@30~";
			    }
		    }
	    }
    }


    //-------------------------------------------------------------------------------
    // ������� ������ �����
    // ---
    bool CreateTolerance3D( ref ITolerance3D tolerance ) 
    {
	    bool res = false;

	    if ( tolerance != null )
	    {
		    // ������� �����������
		    CreateToleranceBranch( ref tolerance );
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
		    // ��������� ���������
		    res = tolerance.Update();
	    }
	    return res;
    }


    //-------------------------------------------------------------------------------
    // �������� ����������� ��� ������� �����
    // ---
    void ToleranceAddBranch( ref ITolerance3D tolerance ) 
    {
	    if ( tolerance != null && doc3D != null )
	    {
			  // ������� � ��������� ������, �� ������� ����� ��������� ����� �����-�������
        ksEntity baseObj = (ksEntity)doc3D.UserSelectEntity( null, "UserFilterProc", 
                                           LoadString("IDS_OBJ"), 0, this );
    		
			  double x = 0, y = 0, z = 0;
			  // ������� ����� ����� �����������
        // ����� ������ ������ �� ������� �������
			  doc3D.UserGetCursor( LoadString("IDS_ENDPOINT"), out x, out y, out z );
    		
			  if ( baseObj != null )
			  {
				  // ������������� ��������� ������� �� API5 � API7
          IModelObject mObj = (IModelObject)kompas.TransferInterface( baseObj, 
                               (int)Kompas6Constants.ksAPITypeEnum.ksAPI7Dual, 0 );
    			
				  if ( mObj != null )
				  {	
					  // �������� ��������� �����������
					  IBranchs3D branchs3D = (IBranchs3D)tolerance;
    				
					  // �������� �����������
					  if ( branchs3D != null )
						  branchs3D.AddBranchByPoint( x, y, z, mObj );

					  // ����������� �� ��������
					  tolerance.set_ArrowType( 1, true );
				  }
			  }
		  }
    }


    //-------------------------------------------------------------------------------
    // ������������� ������ �����
    // ---
    void EditTolerance3D( ref ITolerance3D tolerance ) 
    {
	    if ( tolerance != null )
	    {
		    ToleranceAddBranch( ref tolerance );

		    // ����������������� ���������� ������� �����
		    IToleranceParam tolPar = (IToleranceParam)tolerance;

		    if ( tolPar != null )
		    {
			    // ������ ������� ��������������
			    tolPar.Vertical = true;
			    // ��������� ������� ����� ������������ ������� - ����� ����������
			    tolPar.BasePointPos = Kompas6Constants.ksTablePointEnum.ksTPLeftBottom;
		    }

		    // �������� ��������� ������� ������
		    IFeature7 featObj = (IFeature7)tolerance;
    		
		    if ( featObj != null )
			    // �������� ��� ������� � ������
			    tolerance.Name = GetNewName( ref featObj, "IDS_TOLERANCE" );
    		
		    // �������� ��������� ������� ����� �������
		    IColorParam7 colorPars = (IColorParam7)tolerance;
    		
		    // �������� ���� �� �����-�������
		    if ( colorPars != null )
			    colorPars.Color = RGB( 186, 12, 34 );

		    // ��������� ���������
		    tolerance.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // �������� � �������������� ������� ����� 3D
    // ---
    void Tolerance3DWork()
    {
	    // �������� ��������� ����������� 3D
	    ISymbols3DContainer symbCont = GetSymbols3DContainer();
    	
	    if ( symbCont != null )
	    {
		    // �������� ��������� �������� ����� 3D
		    ITolerances3D tolerCol = symbCont.Tolerances3D;
    		
		    if ( tolerCol != null )
		    {
			    // �������� ����� ������ ����� 3D
			    ITolerance3D newTolerance = tolerCol.Add();
    			
			    if ( newTolerance != null )
			    {
				    // ������� ������ ����� 3D
				    if ( CreateTolerance3D(ref newTolerance) )
				    {	
					    string name = "";
					    // �������� ��������� ������� ������ ����������
					    IFeature7 feature = (IFeature7)newTolerance;
    					
					    // �������� ��� �������
					    if ( feature != null )
						    name = feature.Name;
    					
					    if ( kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
					    {
						    // �������� ������ �� ��������� �� �����
						    ITolerance3D tolerance = tolerCol.get_Tolerance3D( name );
						    // ������������� ������ �����
						    EditTolerance3D( ref tolerance );
					    }
				    }
				    else
					    // ������ ��������� "������ �� ������"
              kompas.ksMessage( LoadString("IDS_NOCREATE") );
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
