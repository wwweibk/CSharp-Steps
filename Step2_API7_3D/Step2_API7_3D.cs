using Kompas6API5;
using KompasAPI7;
using Kompas6Constants;
using Kompas6Constants3D;
using System;
using System.Resources;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Windows.Forms;


namespace Step2_API7_3D
{
  [ClassInterface(ClassInterfaceType.AutoDual)]
	public class Dimensions3D
	{
    private KompasObject      kompas;
    private IApplication      appl;         // ��������� ����������
    private IKompasDocument3D doc;          // ��������� ��������� 3D � API7 
    private ksDocument3D      doc3D;        // ��������� ��������� 3D � API5
    private ResourceManager   resMng = new ResourceManager( typeof(Dimensions3D) );   // �������� ��������
    private int               oType = (int)Kompas6Constants3D.Obj3dType.o3d_edge;

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
        case 1: CreateLineDimension3D();      break;  // �������� ��������� ������� 3D
        case 2: LineDimension3DNavigation();  break;  // ��������� �� ��������� �������� �������� 3D
        case 3: EditLineDimension3D();        break;  // �������������� ��������� ������� 3D
        case 4: RadialDimension3DWork();      break;  // �������� � �������������� ����������� ������� 3D
        case 5: DiametralDimension3DWork();   break;  // �������� � �������������� �������������� ������� 3D
        case 6: AngleDimension3DWork();       break;  // �������� � �������������� �������� ������� 3D
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
    // ���������� ������ �������� � ������� ���������
    // ---
    bool SetLineDimObjectPlane( ref ILineDimension3D dim )
    {
	    bool res = false;
	    // ���������� �������� - �����
	    oType = (int)Kompas6Constants3D.Obj3dType.o3d_edge;
    	
	    if ( doc3D != null )
	    {
		    // ������� � ��������� ������ - �����
        ksEntity obj1 = (ksEntity)doc3D.UserSelectEntity( null, "UserFilterProc", 
                         LoadString("IDS_OBJ1"), 0, this );

		    if ( obj1 != null )
		    {
			    // ��� ���������� - �����
			    oType = (int)Kompas6Constants3D.Obj3dType.o3d_face;
			    // ������� � ��������� ������ - �����
          ksEntity plane = (ksEntity)doc3D.UserSelectEntity( null, "UserFilterProc", 
                            LoadString("IDS_PLANE"), 0, this );

			    if ( plane != null )
			    {
				    // ������������� ���������� ����� � ����� �� API5 � API7
            IModelObject mObj1 = (IModelObject)kompas.TransferInterface( obj1, 
                                 (int)Kompas6Constants.ksAPITypeEnum.ksAPI7Dual, 0 );
            IModelObject mPlane = (IModelObject)kompas.TransferInterface( plane, 
                                  (int)Kompas6Constants.ksAPITypeEnum.ksAPI7Dual, 0 );
    				
				    if ( mObj1 != null && mPlane != null && dim != null )
				    {
					    // ���������� ������ ������ ������� (������ ������ ��������������� 
					    // ������ � ��� ������, ���� ���������� ���������� ����� �������)
					    dim.Object1 = mObj1;
					    // ���������� ������� ��������� �������
					    dim.Plane = mPlane;
					    res = true;
				    }
			    }
		    }
	    }
	    return res;
    }


    //-------------------------------------------------------------------------------
    // ���������� ����� ������� ���������
    // ---
    bool SetNewPlane( ref ILineDimension3D dim )
    {
	    bool res = false;

	    if ( dim != null && doc3D != null )
	    {
		    oType = (int)Kompas6Constants3D.Obj3dType.o3d_face;
		    // ������� � ��������� ������
		    ksEntity obj1 = (ksEntity)doc3D.UserSelectEntity( null, "UserFilterProc", 
                                        LoadString("IDS_NEWPLANE"), 0, this );

		    if ( obj1 != null )
		    {
			    // ������������� ��������� ������� �� API5 � API7
			    IModelObject plane = (IModelObject)kompas.TransferInterface( obj1, 
                               (int)Kompas6Constants.ksAPITypeEnum.ksAPI7Dual, 0 );

			    if ( plane != null )
			    {
				    // ���������� ������� ���������
				    dim.Plane = plane;
				    res = true;
			    }
		    }
	    }
	    return res;
    }
    #endregion


    #region �������� ������ 3D
    //-------------------------------------------------------------------------------
    // �������� ��������� ������� 3D
    // ---
    void CreateLineDimension3D()
    {
	    // �������� ��������� ����������� 3D
	    ISymbols3DContainer symbCont = GetSymbols3DContainer();

	    if ( symbCont != null )
	    {
		    // �������� ��������� �������� �������� 3D
		    ILineDimensions3D dimsCol = symbCont.LineDimensions3D;

		    if ( dimsCol != null )
		    {
			    // �������� ����� �������� ������ 3D
			    ILineDimension3D newDim = (ILineDimension3D)dimsCol.Add( Kompas6Constants3D.
                                     ksObj3dTypeEnum.o3d_lineDimension3D );

			    if ( newDim != null )
			    {
				    bool create = false;

				    // ���� ������� ���������� ������ �������� � ������� ���������
				    if ( SetLineDimObjectPlane(ref newDim) )
				    {
					    // ���������� ����� �������
					    newDim.Length = 30;
					    // ��������� ���������
					    create = (bool)!!newDim.Update();		
    					
				    }
				    else
					    // ������ ��������� "������ �� ������"
              kompas.ksMessage( LoadString("IDS_NOCREATE") );

				    // ���� ������ �� ��������, �������
				    if ( !create )
				    {
					    IFeature7 obj = (IFeature7)newDim;
    					
					    if ( obj != null )
						    obj.Delete();
				    }
			    }
		    }
	    }
    }


    //-------------------------------------------------------------------------------
    // ������ ��������� � ����������� ��������� ������� 3D
    // ---
    void GetLineDimensionPar( ref ILineDimension3D dim )
    {
	    if ( dim != null )
	    {
		    // �������� ����� �������
		    double lenght = dim.Length;
		    double val = 0;

		    // �������� �������� �������
		    IDimensionText dimText = (IDimensionText)dim;

		    if ( dimText != null )
			    val = dimText.NominalValue;

		    // ������������ ���������
        string buf = LoadString( "IDS_LINEDIM3D" ) + "\n\r";
        buf += string.Format( LoadString("IDS_LENGTH"), lenght ) + "\n\r";
        buf += string.Format( LoadString("IDS_DIMVAL"), val );
        kompas.ksMessage( buf );
	    }
    }


    //-------------------------------------------------------------------------------
    // ��������� �� ��������� �������� �������� 3D
    // ---
    void LineDimension3DNavigation()
    {
	    if ( doc != null )
	    {
		    // �������� ��������� ����������� 3D
		    ISymbols3DContainer symbCont = GetSymbols3DContainer();
    		
		    if ( symbCont != null )
		    {
			    // �������� ��������� �������� �������� 3D
			    ILineDimensions3D dimsCol = symbCont.LineDimensions3D;
    			
			    if ( doc3D != null )
			    {
				    // �������� �������� ��������������
				    ksSelectionMng selectMng = (ksSelectionMng)doc3D.GetSelectionMng();

				    if ( dimsCol != null && selectMng != null )
				    {
					    // ���� �� ��������� 
					    for ( long i = 0; i < dimsCol.Count; i++ )
					    {	
						    // �������� ������ �� ��������� �� �������
						    ILineDimension3D lineDim = (ILineDimension3D)dimsCol.get_LineDimension3D( i );

						    if ( lineDim != null )
						    {
							    // ������������� ��������� ������� 3D �� API7 � API5
							    ksEntity dimObj = (ksEntity)kompas.TransferInterface( lineDim,
                                    (int)Kompas6Constants.ksAPITypeEnum.ksAPI5Auto,
                                    (int)Kompas6Constants3D.Obj3dType.o3d_entity );

                  if ( dimObj != null )
							    {
								    // ���������� ������
								    selectMng.Select( dimObj );
								    // ������ ��������� � ����������� �������
								    GetLineDimensionPar( ref lineDim );
								    // ������ ���������
								    selectMng.Unselect( dimObj );
							    }
						    }
					    }
				    }
			    }
		    }
	    }
    }


    //-------------------------------------------------------------------------------
    // �������� ��������� ��������� ������� 3D
    // ---
    void ChangeLineDimensionPar( ref ILineDimension3D dim )
    {
	    if ( dim != null )
	    {
		    // �������� ��������� ���������� �������
		    IDimensionParams dimPars = (IDimensionParams)dim;

		    if ( dimPars != null )
		    {
			    // ��� ������� - �������
			    dimPars.ArrowType1 = Kompas6Constants.ksArrowEnum.ksNotch;
			    dimPars.ArrowType2 = Kompas6Constants.ksArrowEnum.ksNotch;
			    // ������������ ������� - �������
			    dimPars.ArrowPos = Kompas6Constants.ksDimensionArrowPosEnum.ksDimArrowOutside;
			    // ����������� ����� - ������
			    dimPars.ShelfDirection = Kompas6Constants.ksShelfDirectionEnum.ksLSRight;
			    // ����� �����
			    dimPars.ShelfLength = 10;
			    // ���� ������� �����
			    dimPars.ShelfAngle = 45;
		    }
	    }
    }


    //-------------------------------------------------------------------------------
    // �������������� ��������� ������� 3D
    // ---
    void EditLineDimension3D()
    {
	    if ( doc3D != null )
	    {
		    oType = (int)Kompas6Constants3D.Obj3dType.o3d_lineDimension3D;
		    // ������� � ��������� ������
		    ksEntity obj1 = (ksEntity)doc3D.UserSelectEntity( null, "UserFilterProc", 
                                        LoadString("IDS_DIM"), 0, this );

		    if ( obj1 != null )
		    {
			    // ������������� ��������� ������� �� API5 � API7
			    IModelObject mObj1 = (IModelObject)kompas.TransferInterface( obj1, 
                               (int)Kompas6Constants.ksAPITypeEnum.ksAPI7Dual, 0 );

			    if ( mObj1 != null && (int)mObj1.ModelObjectType == (int)Kompas6Constants3D.
                                 Obj3dType.o3d_lineDimension3D )
			    {
				    // �������� ��������� ��������� ������� 3D
				    ILineDimension3D lineDim = (ILineDimension3D)mObj1;

				    if ( lineDim != null )
				    {
					    // ���������� ����� ������� ���������
					    if ( !SetNewPlane(ref lineDim) )
						    // ������ ��������� "�� ������� ���������� ����� ������� ���������"
                kompas.ksMessage( LoadString("IDS_NOTSETPLANE") );

					    // �������� ��������� �������
					    ChangeLineDimensionPar( ref lineDim );
					    // ��������� ����� �������
					    lineDim.Length = lineDim.Length - 10;
					    // ��������� ���������
					    lineDim.Update();
				    }
			    }
			    else
				    // ������ ��������� "������ �� �������� �������� ��������"
            kompas.ksMessage( LoadString("IDS_NOTDIM") );
		    }
	    }
    }
    #endregion


    #region ���������� ������ 3D
    //-------------------------------------------------------------------------------
    // ������� ���������� ������ 3D
    // ---
    bool CreateRadDimension3D( ref IRadialDimension3D dim )
    {
	    bool res = false;

	    if ( dim != null && doc3D != null )
	    {
			  oType = (int)Kompas6Constants3D.Obj3dType.o3d_edge;
			  // ������� � ��������� ������ - �����
        ksEntity edge = (ksEntity)doc3D.UserSelectEntity( null, "UserFilterProc", 
                                        LoadString("IDS_OBJ1"), 0, this );

			  if ( edge != null )
			  {
				  // �������� ��������� ������� �����
				  ksEdgeDefinition edgeDef = (ksEdgeDefinition)edge.GetDefinition();

				  // ���������, �������� �� ����� ��������
				  if ( edgeDef != null && edgeDef.IsCircle() )
				  {
					  // ������������� ��������� ������� �� API5 � API7
					  IModelObject mObj1 = (IModelObject)kompas.TransferInterface( edge, 
                                 (int)Kompas6Constants.ksAPITypeEnum.ksAPI7Dual, 0 );

					  if ( mObj1 != null )
						  // ���������� ������
						  dim.Object1 = mObj1;

					  // �������� ��������� ���������� �������
					  IDimensionParams dimPars = (IDimensionParams)dim;

					  if ( dimPars != null )
					  {
						  // ����������� ����� - �����
						  dimPars.ShelfDirection = Kompas6Constants.ksShelfDirectionEnum.ksLSLeft;
						  // ����� �����
						  dimPars.ShelfLength = 15;
						  // ���� ������� �����
						  dimPars.ShelfAngle = 30;
					  }

					  // ��� ������� - �� �� ������
					  dim.DimensionType = false;
					  // ��������� ���������
					  res = dim.Update();
				  }
				  else
					  // ������ ��������� "����� �� �������� ��������"
            kompas.ksMessage( LoadString("IDS_NOTCIRCLE") );
			  }
		  }
	    return res;
    }


    //-------------------------------------------------------------------------------
    // ������������� ���������� ������ 3D
    // ---
    void EditRadDimension3D( ref IRadialDimension3D dim )
    {
	    if ( dim != null )
	    {
		    // ��� ������� - �� ������
		    dim.DimensionType = true;

		    // �������� ��������� ���������� �������
		    IDimensionParams dimPars = (IDimensionParams)dim;

		    if ( dimPars != null )
			    // ��������� �����
			    dimPars.ShelfDirection = Kompas6Constants.ksShelfDirectionEnum.ksLSNone;

		    // �������� ��������� ������� �������
		    IDimensionText dimText = (IDimensionText)dim;

		    if ( dimText != null )
		    {
			    // ������ ��������
			    dimText.Tolerance = "h6";

			    // �������� ��������
			    dimText.ToleranceOn = true;

			    // �������� ��������� ������ ������� ����������
			    ITextLine lowDev = dimText.LowDeviation;
    			
			    // ������ ��������� ������
			    if ( lowDev != null )
				    lowDev.Str = "+0.021";
		    }
		    // ��������� ���������
		    dim.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // �������� � �������������� ����������� ������� 3D
    // ---
    void RadialDimension3DWork()
    {
	    // �������� ��������� ����������� 3D
	    ISymbols3DContainer symbCont = GetSymbols3DContainer();

	    if ( symbCont != null )
	    {
		    // �������� ��������� ���������� �������� 3D
		    IRadialDimensions3D dimsCol = symbCont.RadialDimensions3D;

		    if ( dimsCol != null )
		    {
			    // �������� ����� ���������� ������ 3D
			    IRadialDimension3D newDim = dimsCol.Add();

			    if ( newDim != null )
			    {
				    // ������� ���������� ������ 3D
				    if ( CreateRadDimension3D(ref newDim) )
				    {	
					    string name = "";
					    // �������� ��������� ������� ������ ����������
					    IFeature7 feature = (IFeature7)newDim;
              				
					    // �������� ��� �������
					    if ( feature != null )
						    name = feature.Name;

					    if ( kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
					    {
						    // �������� ������ �� ��������� �� �����
						    IRadialDimension3D radDim = dimsCol.get_RadialDimension3D( name );
						    // ������������� ������
						    EditRadDimension3D( ref radDim );
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


    #region ������������� ������ 3D
    //-------------------------------------------------------------------------------
    // ������� �������������� ������
    // ---
    bool CreateDiamDimension3D( ref IDiametralDimension3D dim )
    {
	    bool res = false;
    	
	    if ( dim != null && doc3D != null )
	    {
			  oType = (int)Kompas6Constants3D.Obj3dType.o3d_edge;
			  // ������� � ��������� ������ - �����
        ksEntity edge = (ksEntity)doc3D.UserSelectEntity( null, "UserFilterProc", 
                                        LoadString("IDS_OBJ1"), 0, this );
			  
        if ( edge != null )
			  {
				  // �������� ��������� ������� �����
				  ksEdgeDefinition edgeDef = (ksEdgeDefinition)edge.GetDefinition();
    			
				  // ���������, �������� �� ����� ��������
				  if ( edgeDef != null && edgeDef.IsCircle() )
				  {
					  // ������������� ��������� ������� �� API5 � API7
            IModelObject mObj1 = (IModelObject)kompas.TransferInterface( edge, 
                                 (int)Kompas6Constants.ksAPITypeEnum.ksAPI7Dual, 0 );
    				
					  if ( mObj1 != null )
						  // ���������� ������
						  dim.Object1 = mObj1;
    				
					  // ��������� ���������
					  res = dim.Update();
				  }
				  else
					  // ������ ��������� "����� �� �������� ��������"
            kompas.ksMessage( LoadString("IDS_NOTCIRCLE") );
			  }
		  }
	    return res;
    }


    //-------------------------------------------------------------------------------
    // ������������� ������������� ������
    // ---
    void EditDiamDimension3D( ref IDiametralDimension3D dim )
    {
	    if ( dim != null )
	    {
		    // ��� ������� - � �������
		    dim.DimensionType = true;

		    // �������� ��������� ������� �������
		    IDimensionText dimText = (IDimensionText)dim;

		    if ( dimText != null )
		    {
			    // �������� ��������� ������ ��� ��������� ��������
			    IText txtUnder = dimText.TextUnder;
    			
			    // ������ ��������� ������
			    if ( txtUnder != null )
				    txtUnder.Str = LoadString( "IDS_DIMTEXT" );

			    // ������������ �����
			    dimText.Underline = true;
		    }
		    // ��������� ���������
		    dim.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // �������� � �������������� �������������� ������� 3D
    // ---
    void DiametralDimension3DWork()
    {
	    // �������� ��������� ����������� 3D
	    ISymbols3DContainer symbCont = GetSymbols3DContainer();
    	
	    if ( symbCont != null )
	    {
		    // �������� ��������� ������������� �������� 3D
		    IDiametralDimensions3D dimsCol = symbCont.DiametralDimensions3D;
    		
		    if ( dimsCol != null )
		    {
			    // �������� ����� ������������� ������ 3D
			    IDiametralDimension3D newDim = dimsCol.Add();
    			
			    if ( newDim != null )
			    {
				    // ������� ������������� ������ 3D
				    if ( CreateDiamDimension3D(ref newDim) )
				    {	
					    string name = "";
					    // �������� ��������� ������� ������ ����������
					    IFeature7 feature = (IFeature7)newDim;
    					
					    // �������� ��� �������
					    if ( feature != null )
						    name = feature.Name;
    					
					    if ( kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
					    {
						    // �������� ������ �� ��������� �� �����
						    IDiametralDimension3D diamDim = dimsCol.get_DiametralDimension3D( name );
						    // ������������� ������
						    EditDiamDimension3D( ref diamDim );
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


    #region ������� ������ 3D
    //-------------------------------------------------------------------------------
    // ������� ������� ������ 3D
    // ---
    bool CreateAngleDimension3D( ref IAngleDimension3D dim )
    {
	    bool res = false;

	    if ( dim != null && doc3D != null )
	    {
			  // ������� � ��������� 1-� ������
        ksEntity obj1 = (ksEntity)doc3D.UserSelectEntity( null, "UserFilterProc", 
                                        LoadString("IDS_OBJECT1"), 0, this );
			  // ������� � ��������� 2-� ������
        ksEntity obj2 = (ksEntity)doc3D.UserSelectEntity( null, "UserFilterProc", 
                                        LoadString("IDS_OBJECT2"), 0, this );

			  if ( obj1 != null && obj2 != null )
			  {
				  // ������������� ���������� �������� �� API5 � API7
          IModelObject mObj1 = (IModelObject)kompas.TransferInterface( obj1, 
                               (int)Kompas6Constants.ksAPITypeEnum.ksAPI7Dual, 0 );
          IModelObject mObj2 = (IModelObject)kompas.TransferInterface( obj2, 
                               (int)Kompas6Constants.ksAPITypeEnum.ksAPI7Dual, 0 );

				  if ( mObj1 != null && mObj2 != null )
				  {
					  // 1-� ������
					  dim.Object1 = mObj1;
					  // 2-� ������
					  dim.Object2 = mObj2;
					  // ����� ��������� �����
					  dim.Length = 20;
					  // ��������� ���������
					  res = dim.Update();
				  }
			  }
		  }
	    return res;
    }


    //-------------------------------------------------------------------------------
    // ������������� ������� ������ 3D
    // ---
    void EditAngleDimension3D( ref IAngleDimension3D dim )
    {
	    if ( dim != null )
	    {
		    // ��� ������� - �� ������������ (�����) ����
		    dim.DimensionType = Kompas6Constants.ksAngleDimTypeEnum.ksADMaxAngle;
		    // ��������� ����� ��������� �����
		    dim.Length = dim.Length + 10;

		    // �������� ��������� ���������� �������
		    IDimensionParams dimPars = (IDimensionParams)dim;

		    if ( dimPars != null )
			    // ���������� ��������� ������� - �����������, � ������� �����
			    dimPars.TextOnLine = Kompas6Constants.ksDimensionTextPosEnum.
                               ksDimTextParallelInCut;

		    // �������� ��������� ������� �������
		    IDimensionText dimText = (IDimensionText)dim;

		    if ( dimText != null )
			    // �������� ������ ����������� ������ �� ����������
			    dimText.TextFormat = Kompas6Constants.ksDimTextFormatEnum.ksDimTextFormatGDD;

		    // ��������� ���������
		    dim.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // �������� � �������������� �������� ������� 3D
    // ---
    void AngleDimension3DWork()
    {
	    // �������� ��������� ����������� 3D
	    ISymbols3DContainer symbCont = GetSymbols3DContainer();
    	
	    if ( symbCont != null )
	    {
		    // �������� ��������� ������� �������� 3D
		    IAngleDimensions3D dimsCol = symbCont.AngleDimensions3D;
    		
		    if ( dimsCol != null )
		    {
			    // �������� ����� ������� ������ 3D
			    IAngleDimension3D newDim = dimsCol.Add();
    			
			    if ( newDim != null )
			    {
				    // ������� ������� ������ 3D
				    if ( CreateAngleDimension3D(ref newDim) )
				    {	
					    string name = "";
					    // �������� ��������� ������� ������ ����������
					    IFeature7 feature = (IFeature7)newDim;
    					
					    // �������� ��� �������
					    if ( feature != null )
						    name = feature.Name;
    					
					    if ( kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
					    {
						    // �������� ������ �� ��������� �� �����
						    IAngleDimension3D angDim = dimsCol.get_AngleDimension3D( name );
						    // ������������� ������
						    EditAngleDimension3D( ref angDim );
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
