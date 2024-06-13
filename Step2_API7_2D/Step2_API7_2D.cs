using Kompas6API5;
using KompasAPI7;
using Kompas6Constants;
using KAPITypes;
using System;
using System.Resources;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Windows.Forms;


namespace Step2_API7_2D
{
  //-------------------------------------------------------------------------------
  // ������� ������ � ��������� 2D � API7
  // ---
  [ClassInterface(ClassInterfaceType.AutoDual)]
  public class SamplesDimension
  {
    private KompasObject      kompas;
    private IApplication      appl;         // ��������� ����������
    private IKompasDocument2D doc;          // ��������� ��������� 2D � API7 
    private ksDocument2D      doc2D;        // ��������� ��������� 2D � API5
    private ResourceManager   resMng = new ResourceManager( typeof(SamplesDimension) );       // �������� ��������
   
    
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
          result = LoadString( "IDS_MENU7" );
          command = 7;
        break;

        case 8:
          result = LoadString( "IDS_MENU8" );
          command = 8;
        break;

        case 9:
          result = LoadString( "IDS_MENU9" );
          command = 9;
        break;

        case 10:
          result = LoadString( "IDS_MENU10" );
          command = 10;
        break;

        case 11:
          result = LoadString( "IDS_MENU11" );
          command = 11;
        break;

        case 12:
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
        case 1:   CreateLineDimension();      break;  // ������� �������� ������
        case 2:   LineDimensionNavigation();  break;  // ��������� �� ��������� �������� ��������
        case 3:   EditLineDimension();        break;  // �������������� ��������� �������
        case 4:   RadialDimensionWork();      break;  // �������� � �������������� ����������� �������
        case 5:   DiamrtralDimensionWork();   break;  // �������� � �������������� �������������� �������
        case 6:   AngleDimensionWork();       break;  // �������� � �������������� �������� �������
        case 7:   ArcDimensionWork();         break;  // �������� � �������������� ������� ���� ����������
        case 8:   BreakLineDimensionWork();   break;  // �������� � �������������� ��������� ������� � �������
        case 9:   BreakRadialDimensionWork(); break;  // �������� � �������������� ����������� ������� � �������
        case 10:  BreakAngleDimensionWork();  break;  // �������� � �������������� �������� ������� � �������
        case 11:  HeightDimensionWork();      break;  // �������� � �������������� ������� ������
      }
    }


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

    #region �������� ������
    //-------------------------------------------------------------------------------
    // ������� �������� ������
    // ---
    void CreateLineDimension()
    {
      // �������� ��������� �������� �����������
      ISymbols2DContainer symbCont = GetSymbols2DContainer();

      if ( symbCont != null )
      {
        // �������� ��������� �������� ��������
        ILineDimensions dimCol = symbCont.LineDimensions;

        if ( dimCol != null )
        {
          //�������� ������ � ���������
          ILineDimension newDim = dimCol.Add();

          if ( newDim != null )
          {
            // ���������� ������ ����� �������� �������
            newDim.X1 = 50;
            newDim.Y1 = 150;
            // ���������� ������ ����� �������� �������
            newDim.X2 = 100;
            newDim.Y2 = 150;
            // ��������� ��������� �����
            newDim.X3 = 75;
            newDim.Y3 = 180;
            // ��� ���������� ��������� �������
            newDim.Orientation = Kompas6Constants.ksLineDimensionOrientationEnum.ksLinDHorizontal;
            // ��������� ���������
            newDim.Update();    
          }
        }
      }
    }


    //-------------------------------------------------------------------------------
    // ������ ��������� � ����������� ��������� �������
    // ---
    void GetLineDimensionParam( ILineDimension dim )
    {
      if ( dim != null )
      {
        // ���������� ������ ����� �������� �������
        double x1 = dim.X1;
        double y1 = dim.Y1;
        // ���������� ������ ����� �������� �������
        double x2 = dim.X2;
        double y2 = dim.Y2;
        // ������ ��������� � ����������� �������
        string buf = string.Format( LoadString("IDS_COORDS"), x1, y1, x2, y2 );
        kompas.ksMessage( buf );
      }
    }


    //-------------------------------------------------------------------------------
    // ��������� �� ��������� �������� ��������
    // ---
    void LineDimensionNavigation()
    {
      // �������� ��������� �������� �����������
      ISymbols2DContainer symbCont = GetSymbols2DContainer();

      if ( symbCont != null )
      {
        // �������� ��������� �������� ��������
        ILineDimensions dimCol = symbCont.LineDimensions;

        if ( dimCol != null )
        {
          for ( long i = 0; i < dimCol.Count; i++ )
          {
            // �������� ������ �� ��������� �� �������
            ILineDimension lineDim = dimCol.get_LineDimension( i );

            if ( lineDim != null )
            {
              int dimRef = lineDim.Reference;

              if ( doc2D != null )
              {
                // ���������� ������
                doc2D.ksLightObj( dimRef, 1/*�������� ���������*/ );
                // ������ ��������� � ����������� ��������� �������
                GetLineDimensionParam( lineDim );
                // �������� ������
                doc2D.ksLightObj( dimRef, 0/*��������� ���������*/ );
              }
            }
          }
        }
      }
    }


    //-------------------------------------------------------------------------------
    // �������� ��������� � ��������� �������
    // ---
    void ChangeLineDimensionParam( ref ILineDimension dim )
    {
      if ( dim != null )
      {
        // �������� ��������� ���������� �������
        IDimensionParams dimPar = (IDimensionParams)dim;
    		
        if ( dimPar != null )
        {
          // ����������� ����� - �����
          dimPar.ShelfDirection = Kompas6Constants.ksShelfDirectionEnum.ksLSLeft;
          // ���� ������� - ���� 90 ����.
          dimPar.ArrowType1 = Kompas6Constants.ksArrowEnum.ksRightAngle;
          dimPar.ArrowType2 = Kompas6Constants.ksArrowEnum.ksRightAngle;
        }
      		
        // ���������� ����� ������ �����
        dim.ShelfX = 60;
        dim.ShelfY = 200;
      }
    }


    //-------------------------------------------------------------------------------
    // �������� ����� � ��������� �������
    // ---
    void ChangeLineDimensionText( ref ILineDimension dim )
    {
      if ( dim != null )
      {
        // �������� ��������� ������ �������
        IDimensionText dimText = (IDimensionText)dim;
    		
        if ( dimText != null )
        {
          // ����� � �����
          dimText.Rectangle = true;
          // ����������� �����
          dimText.Underline = true;
      			
          // �������� ��������� ������ �� �������� �������
          ITextLine prefix = dimText.Prefix;
      			
          // �������� �����
          if ( prefix != null )
            prefix.Str = LoadString( "IDS_DIM" );
        }
      }
    }


    //-------------------------------------------------------------------------------
    // �������� �������� ������
    // ---
    void LineDimChange( ref ILineDimension dim )
    {
      if ( dim != null )
      {
        // �������� ��������� �������
        ChangeLineDimensionParam( ref dim );
        // �������� ����� �������
        ChangeLineDimensionText( ref dim );
        // ��������� ���������
        dim.Update();
      }
    }


    //-------------------------------------------------------------------------------
    // �������������� ��������� �������
    // ---
    void EditLineDimension()
    {
      // ��������� ���������� ������� � �������
      ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct( 
                           (int)Kompas6Constants.StructType2DEnum.ko_RequestInfo );
      info.Init();
      info.commandsString = LoadString( "IDS_COMMAND1" );
      double x = 0, y = 0;
  	
      if ( doc2D != null && info != null )
      {
        // ������� ������ � ���������
        if ( doc2D.ksCursorEx(info, ref x, ref y, null, null) != 0 )
        {
          // ����� ������ �� ��������� �����������
          int pObj = doc2D.ksFindObj( x, y, doc2D.ksGetCursorLimit() );
    		
          // ���� ��� �������� ������, �������������
          if ( doc2D.ksGetObjParam(pObj, 0, 0) == ldefin2d.LDIMENSION_OBJ/*�������� ������*/ )
          {
            // �������� ��������� ��������� ������� �� ��������� �������
            ILineDimension lineDim = (ILineDimension)kompas.TransferReference( pObj, 
                                      doc2D.reference );
            // �������� �������� ������
            LineDimChange( ref lineDim );
          }
          else
            kompas.ksMessage( LoadString( "IDS_NOTDIM") );
        }
      } 
    }
    #endregion


    #region ���������� ������
    //-------------------------------------------------------------------------------
    // ������� ���������� ������
    // ---
    bool CreateRadDimension( ref IRadialDimension dim )
    {
      bool res = false;

      if ( dim != null )
      {
        // ��������� ���������� ������� � �������
        ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct( 
                             (int)Kompas6Constants.StructType2DEnum.ko_RequestInfo );
        info.Init();
        info.commandsString = LoadString( "IDS_COMMAND2" );
        double x = 0, y = 0;
    		
        if ( doc2D != null && info != null )
        {
          // ������� ������ � ���������
          if ( doc2D.ksCursorEx( info, ref x, ref y, null, null) != 0 )
          {
            // ����� ������ �� ��������� �����������
            int pObj = doc2D.ksFindObj( x, y, doc2D.ksGetCursorLimit() );
            // ��������� ���������� ����������
            CircleParam par = (CircleParam)kompas.GetParamStruct( 
                              (int)Kompas6Constants.StructType2DEnum.ko_CircleParam );
            par.Init();

            // �������� �� ������ �����������
            if ( doc2D.ksGetObjParam(pObj, par, ldefin2d.ALLPARAM ) == ldefin2d.CIRCLE_OBJ )
            {
              // �������� ��������� ������������ ������� �� ���������
              IDrawingObject circle = (IDrawingObject)kompas.TransferReference( pObj, doc2D.reference );
    				
              if ( circle != null )
              {
                // ������ ������� ������ �������
                dim.BaseObject = circle;
                // ������ ���������� ������ �������, ����������� � ������� ��������� ����������
                dim.Xc = par.xc;
                dim.Yc = par.yc;
                // ������ ������ �������, ����������� � �������� ��������� ����������
                dim.Radius = par.rad;
                // ��� ������� - �� ������
                dim.DimensionType = true;
                // ���� ������� ��������� �����
                dim.Angle = 30;
                // ��������� ���������
                res = dim.Update();
              }
              else
                kompas.ksMessage( LoadString("IDS_NOOBJ") );
            }
            else
              kompas.ksMessage( LoadString("IDS_NOTCREATE1") );
          } 
        }
      }
      return res;
    }


    //-------------------------------------------------------------------------------
    // �������� ��������� � ����� � ����������� �������
    // ---
    void ChangeRadialDimensionParamText( ref IRadialDimension dim )
    {
      if ( dim != null )
      {
        // �������� ��������� ���������� �������
        IDimensionParams dimPar = (IDimensionParams)dim;

        if ( dimPar != null )
        {
          // ����������� ����� - ������
          dimPar.ShelfDirection = Kompas6Constants.ksShelfDirectionEnum.ksLSRight;
          // ���� ������� �����
          dimPar.ShelfAngle = 180;
          // ����� �����
          dimPar.ShelfLength = 30;
          // ��� ������� - �������
          dimPar.ArrowType1 = Kompas6Constants.ksArrowEnum.ksNotch;
        }

        // �������� ��������� ������� �������
        IDimensionText dimText = (IDimensionText)dim;

        if ( dimText != null )
        {
          // ������ ����� ��������� - ������
          dimText.Sign = 3;

          // �������� ��������� ������ ������� ���������
          ITextLine unit = dimText.Unit;

          // ������ ��������� ������
          if ( unit != null )
            unit.Str = LoadString( "IDS_UNIT" );
        }
      }
    }


    //-------------------------------------------------------------------------------
    // ������������� ���������� ������
    // ---
    void EditRadialDimension( ref IRadialDimension dim )
    {
      if ( dim != null )
      {
        // ��� ������� - �� �� ������
        dim.DimensionType = false;
        // �������� ��������� � ����� �������
        ChangeRadialDimensionParamText( ref dim );
        // ��������� ���������
        dim.Update();
      }
    }


    //-------------------------------------------------------------------------------
    // �������� � �������������� ����������� �������
    // ---
    void RadialDimensionWork()
    {
      // �������� ��������� �������� �����������
      ISymbols2DContainer symbCont = GetSymbols2DContainer();

      if ( symbCont != null )
      {
        // �������� ��������� ���������� ��������
        IRadialDimensions radCol = symbCont.RadialDimensions;

        if ( radCol != null )
        {
          // �������� ���������� ������ � ���������
          IRadialDimension radDim = radCol.Add();

          if ( radDim != null )
          {
            // ������� ���������� ������
            bool create = CreateRadDimension( ref radDim );
            // �������� �������� �������
            int dimRef = radDim.Reference;

            if ( create && kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
            {
              // �������� ������ �� ��������� �� ���������
              IRadialDimension radDimension = radCol.get_RadialDimension(dimRef);
              // ������������� ���������� ������
              EditRadialDimension( ref radDimension );
            }
          }
        }
      }
    }
    #endregion


    #region ������������� ������
    //-------------------------------------------------------------------------------
    // ������� ������������� ������
    // ---
    bool CreateDiamrtralDimension( ref IDiametralDimension dim )
    {
      bool res = false;

      if ( dim != null )
      {
        // ��������� ���������� ������� � �������
        ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct( 
                             (int)Kompas6Constants.StructType2DEnum.ko_RequestInfo );
        info.Init();
        info.commandsString = LoadString( "IDS_COMMAND2" );
        double x = 0, y = 0;
    		
        // ������� ������ � ���������
        if ( doc2D.ksCursorEx( info, ref x, ref y, null, null) != 0 )
        {
          // ����� ������ �� ��������� �����������
          int pObj = doc2D.ksFindObj( x, y, doc2D.ksGetCursorLimit() );
          // ��������� ���������� ����������
          CircleParam par = (CircleParam)kompas.GetParamStruct( 
                            (int)Kompas6Constants.StructType2DEnum.ko_CircleParam );
          par.Init();
    			
          // �������� �� ������ �����������
          if ( doc2D.ksGetObjParam(pObj, par, ldefin2d.ALLPARAM ) == ldefin2d.CIRCLE_OBJ )
          {
            // �������� ��������� ������������ ������� �� ���������
            IDrawingObject circle = (IDrawingObject)kompas.TransferReference( pObj, 
                                     doc2D.reference );
    				
            if ( circle != null )
            {
              // ������ ������� ������ �������
              dim.BaseObject = circle;
              // ������ ���������� ������ �������, ����������� � ������� ��������� ����������
              dim.Xc = par.xc;
              dim.Yc = par.yc;
              // ������ ������ �������, ����������� � �������� ��������� ����������
              dim.Radius = par.rad;
              // ��� ������� - ������ ��������� �����
              dim.DimensionType = true;
              // ���� ������� ��������� �����
              dim.Angle = 45;
              // ��������� ���������
              res = dim.Update();
            }
            else
              kompas.ksMessage( LoadString("IDS_NOOBJ") );
          }
          else
            kompas.ksMessage( LoadString("IDS_NOTCREATE2") );
        }
      }
      return res;
    }


    //-------------------------------------------------------------------------------
    // �������� ��������� � ����� � �������������� �������
    // ---
    void ChangeDiametralDimensionParamText( ref IDiametralDimension dim )
    {
      if ( dim != null )
      {
        // �������� ��������� ���������� �������
        IDimensionParams dimPar = (IDimensionParams)dim;
    		
        if ( dimPar != null )
        {
          // ����������� ����� - ������
          dimPar.ShelfDirection = Kompas6Constants.ksShelfDirectionEnum.ksLSRight;
          // ����� �����
          dimPar.ShelfLength = 20;
          // ��� ������� - �����
          dimPar.ArrowType1 = Kompas6Constants.ksArrowEnum.ksPoint;
          dimPar.ArrowType2 = Kompas6Constants.ksArrowEnum.ksPoint;
        }
    		
        // �������� ��������� ������� �������
        IDimensionText dimText = (IDimensionText)dim;
    		
        if ( dimText != null )
        {
          // ������ ����� ��������� - �������
          dimText.Sign = 1;
          // �����������
          dimText.Underline = true;
        			
          // �������� ��������� ������ ������� ���������
          ITextLine suffix = dimText.Suffix;
        			
          // ������ ��������� ������
          if ( suffix != null )
            suffix.Str = LoadString( "IDS_UNIT" );
        }
      }
    }


    //-------------------------------------------------------------------------------
    // ������������� ������������� ������
    // ---
    void EditDiametralDimension( ref IDiametralDimension dim )
    {
      if ( dim != null )
      {
        // ��� ������� - ��������� ����� � �������
        dim.DimensionType = false;
        // ���� ������� ��������� �����
        dim.Angle = 90;
        // �������� ��������� � ����� �������
        ChangeDiametralDimensionParamText( ref dim );
        // ��������� ���������
        dim.Update();
      }
    }

   
    //-------------------------------------------------------------------------------
    // �������� � �������������� �������������� �������
    // ---
    void DiamrtralDimensionWork()
    {
      // �������� ��������� �������� �����������
      ISymbols2DContainer symbCont = GetSymbols2DContainer();

      if ( symbCont != null )
      {
        // �������� ��������� ������������� ��������
        IDiametralDimensions diamCol = symbCont.DiametralDimensions;
    		
        if ( diamCol != null )
        {
          // �������� ������������� ������ � ���������
          IDiametralDimension diamDim = diamCol.Add();
    			
          if ( diamDim != null )
          {
            // ������� ������������� ������
            bool create = CreateDiamrtralDimension( ref diamDim );
            // �������� �������� �������
            int dimRef = diamDim.Reference;
    				
            if ( create && kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
            {
              // �������� ������ �� ��������� �� ���������
              IDiametralDimension diamDimension = diamCol.get_DiametralDimension( dimRef );
              // ������������� ������������� ������
              EditDiametralDimension( ref diamDimension );
            }
          }
        }
      }
    }
    #endregion


    #region ������� ������
    //-------------------------------------------------------------------------------
    // ������� ������� ������
    // ---
    bool CreateAngleDimension( ref IAngleDimension dim )
    {
      bool res = false;

      if ( dim != null && doc2D != null )
      {
        // ���������� ������� �� ��������, ��������� � �������
        // ��������� ���������� ������� � �������
        /*ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct( (int)Kompas6Constants.StructType2DEnum.ko_RequestInfo );
        info.Init();
        info.commandsString = LoadString( "IDS_COMMAND3" );
        double x = 0, y = 0;
	
        // ������� ������ ������� ������ � ���������
        if ( doc2D.ksCursorEx(info, ref x, ref y, null, null) != 0 )
        {
          // ����� ������ �� ��������� �����������
          int pObj1 = doc2D.ksFindObj( x, y, 1 );
          // �������� ��� ������� �������
          int type1 = doc2D.ksGetObjParam( pObj1, 0, 0 );

          if ( type1 == ldefin2d.LINESEG_OBJ || type1 == ldefin2d.POLYLINE_OBJ || type1 == ldefin2d.RECTANGLE_OBJ )
          {
            // ���������� ��������� ������
            doc2D.ksLightObj( pObj1, 1 );
            info.commandsString = LoadString( "IDS_COMMAND4" );
			
            // ������� ������ ������� ������ � ���������
            if ( doc2D.ksCursorEx(info, ref x, ref y, null, null) != 0 )
            {
              // ����� ������ �� ��������� �����������
              int pObj2 = doc2D.ksFindObj( x, y, 1 );
              // �������� ��� ������� �������
              int type2 = doc2D.ksGetObjParam( pObj2, 0, 0 );

              if ( type1 == ldefin2d.LINESEG_OBJ || type1 == ldefin2d.POLYLINE_OBJ || type1 == ldefin2d.RECTANGLE_OBJ )
              {
                // ���������� ��������� ������
                doc2D.ksLightObj( pObj2, 1 );
                // �������� ��������� ������������ ������� �� ���������
                IDrawingObject baseObj1 = (IDrawingObject)kompas.TransferReference( pObj1, doc2D.reference );
                IDrawingObject baseObj2 = (IDrawingObject)kompas.TransferReference( pObj2, doc2D.reference );

                if ( baseObj1 != null && baseObj2 != null )
                {
                  // ������ ������� ������� �������
                  dim.BaseObject1 = baseObj1;
                  dim.BaseObject2 = baseObj2;
                  // ������ ����
                  dim.Radius = 20;
                  // ��������� ���������
                  res = dim.Update();
                  // ������ ��������� �������
                  doc2D.ksLightObj( pObj2, 0 );
                }
                else
                  kompas.ksMessage( LoadString("IDS_NOOBJS") );
					
                // ������ ��������� �������
                doc2D.ksLightObj( pObj1, 0 );
              }
              else
                kompas.ksMessage( LoadString("IDS_NOTCREATE3") );
            }
          }
          else
            kompas.ksMessage( LoadString("IDS_NOTCREATE3") );
        }*/
        // ���������� ������� �� ���� ��������		
        // ��������� ������ ������� ������
        int line1 = doc2D.ksLineSeg( 80, 120, 100, 120, 1 );
        IDrawingObject baseObj1 = (IDrawingObject)kompas.TransferReference( line1, 
                                   doc2D.reference );
        // ��������� ������ ������� ������
        int line2 = doc2D.ksLineSeg( 80, 120, 120, 160, 1 );
        IDrawingObject baseObj2 = (IDrawingObject)kompas.TransferReference( line2, 
                                   doc2D.reference );
		
        if ( baseObj1 != null && baseObj2 != null )
        {
          // ������ ������� �������
          dim.BaseObject1 = baseObj1;
          dim.BaseObject2 = baseObj2;
          // ���������� ������
          dim.Xc = 90;
          dim.Yc = 150;
          // ������ ����
          dim.Radius = 25;
          // ��������� ����  ��������� ����
          dim.Angle1 = 45;
          // �������� ����  ��������� ����
          dim.Angle2 = 30;
          // ���������� ����� ������ ������ �������� �����
          dim.X1 = 80;
          dim.Y1 = 120;
          // ���������� ����� ������ ������ �������� �����
          dim.X2 = 100;
          dim.Y2 = 160;
          // ��� ������� - �� ����������� (������) ����
          dim.DimensionType = Kompas6Constants.ksAngleDimTypeEnum.ksADMinAngle;
			
          // �������� �����
          IDimensionParams dimPar = (IDimensionParams)dim;
			
          if ( dimPar != null )
            // ����������� ����� - ������
            dimPar.ShelfDirection = Kompas6Constants.ksShelfDirectionEnum.ksLSRight;
					
          // ����� ������ �����
          dim.ShelfX = 160;
          dim.ShelfY = 140;
          // ����������� ��������� ���� - ������ ������� �������
          dim.Direction = false;
          // ������ ����� ��� ����� �� ����
          dim.X3 = 140;
          dim.Y3 = 120;
          // ��������� ���������
          res = dim.Update();
        }
      }
      return res;
    }


    //-------------------------------------------------------------------------------
    // �������������� �������� �������
    // ---
    void EditAngleDimension( ref IAngleDimension dim )
    {
      if ( dim != null )
      {
        // ��� ������� - �� ������������ (�����) ����
        dim.DimensionType = Kompas6Constants.ksAngleDimTypeEnum.ksADMaxAngle;
        // ����������� ��������� ����  - �� ������� �������
        dim.Direction = true;
        // ����� ������ �����
        dim.ShelfX = 150;
        dim.ShelfY = 170;
        // ������ ����
        dim.Radius = 40;
        // ��������� ���������
        dim.Update();
      }
    }


    //-------------------------------------------------------------------------------
    // �������� � �������������� �������� �������
    // ---
    void AngleDimensionWork()
    {
	    // �������� ��������� �������� �����������
	    ISymbols2DContainer symbCont = GetSymbols2DContainer();
    	
	    if ( symbCont != null )
	    {
		    // �������� ��������� ������� ��������
		    IAngleDimensions angCol = symbCont.AngleDimensions;
    		
		    if ( angCol != null )
		    {
			    // �������� ������� ������ � ���������
			    IAngleDimension angDim = angCol.Add( Kompas6Constants.DrawingObjectTypeEnum.
                                               ksDrADimension );
    			
			    if ( angDim != null )
			    {
				    // ������� ������� ������
				    bool create = CreateAngleDimension( ref angDim );
				    // �������� �������� �������
				    int dimRef = angDim.Reference;
    				
            if ( create && kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
				    {
					    // �������� ������ �� ��������� �� ���������
					    IAngleDimension angDimension = angCol.get_AngleDimension( dimRef );
					    // ������������� ������� ������
					    EditAngleDimension( ref angDimension );
				    }
			    }
		    }
	    }
    }
    #endregion


    #region ������ ���� ����������
    //-------------------------------------------------------------------------------
    // �������� ������� ���� ����������
    // ---
    bool CreateArcDimension( ref IArcDimension dim )
    {
      bool res = false;

	    if ( dim != null && doc2D != null )
	    {
		    // ��������� ���������� ������� � �������
        ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct( 
                             (int)Kompas6Constants.StructType2DEnum.ko_RequestInfo );
		    info.Init();
		    info.commandsString = LoadString( "IDS_COMMAND5" );
		    double x = 0, y = 0;
    		
		    // ������� ������ � ���������
        if ( doc2D.ksCursorEx(info, ref x, ref y, null, null) != 0 )
		    {
			    // ����� ������ �� ��������� �����������
          int pObj = doc2D.ksFindObj( x, y, 1 );
			    // ���������� ��������� ������
			    doc2D.ksLightObj( pObj, 1 );
			    // ��������� ���������� ����
			    ArcByPointParam par = (ArcByPointParam)kompas.GetParamStruct( 
                                (int)Kompas6Constants.StructType2DEnum.ko_ArcByPointParam );
    			par.Init();

			    // �������� �� ������ �����
          if ( doc2D.ksGetObjParam(pObj, par, ldefin2d.POINT_ARC_PARAM ) == ldefin2d.ARC_OBJ )
			    {
				    // �������� ��������� ������������ ������� �� ���������
				    IDrawingObject arc = (IDrawingObject)kompas.TransferReference( pObj, 
                                  doc2D.reference );
    				
				    if ( arc != null )
				    {
					    // ������� ������
					    dim.BaseObject = arc;
					    // ���������� ������
					    dim.Xc = par.xc;
					    dim.Yc = par.yc;
					    // ���������� ������ ����� ����
					    dim.X1 = par.x1;
					    dim.Y1 = par.y1;
					    // ���������� ������ ����� ����
					    dim.X2 = par.x2;
					    dim.Y2 = par.y2;
    					
					    info.commandsString = LoadString( "IDS_COMMAND6" );					
					    // ������� ��������� ��������� �����
              if ( doc2D.ksCursorEx(info, ref x, ref y, null, null) != 0 )
					    {
						    dim.X3 = x;
						    dim.Y3 = y;
					    }

					    // ����������� ��������� ����
					    dim.Direction = true;
					    // ��� ������� - ������������ �������� �����
					    dim.DimensionType = false;
					    // ��������� �� ������ � ����
					    dim.TextPointer = true;
					    // ��������� ���������
					    res = dim.Update();
				    }
				    else
              kompas.ksMessage( LoadString("IDS_NOOBJ") );
			    }
			    else
            kompas.ksMessage( LoadString("IDS_NOTCREATE4") );

			    // ������ ��������� �������
          doc2D.ksLightObj( pObj, 0 );
		    }
	    }
	    return res;
    }


    //-------------------------------------------------------------------------------
    // �������������� ������� ���� ����������
    // ---
    void EditArcDimension( ref IArcDimension dim )
    {
      if ( dim != null )
      {
        // ��� ������� - �������� ����� �� ������
        dim.DimensionType = true;
        // ��������� �� ������ � ����
        dim.TextPointer = false;
        // ����������� ��������� ���� - �������� �� ���������������
        dim.Direction = !dim.Direction;

        // �������� ��������� ���������� �������
        IDimensionParams dimPar = (IDimensionParams)dim;

        if ( dimPar != null )
        {
          // ��� �������
          dimPar.ArrowType1 = Kompas6Constants.ksArrowEnum.ksLeaderPoint;
          dimPar.ArrowType2 = Kompas6Constants.ksArrowEnum.ksLeaderPoint;
          // ����������� ����� - ����
          dimPar.ShelfDirection = Kompas6Constants.ksShelfDirectionEnum.ksLSDown;
          // ���� ������� �����
          dimPar.ShelfAngle = 180;
        }

        dim.Update();

        // �������� ��������� ������ �������
        IDimensionText dimText = (IDimensionText)dim;

        if ( dimText != null )
        {
          // ��������� �������������� ����������� ������������ ��������
          dimText.AutoNominalValue = false;
          // ���������� ����� ����������� ��������
          dimText.NominalValue = 50;
        			
          // �������� ��������� ������ ������� ���������
          ITextLine unit = dimText.Unit;
        			
          // ������ ��������� ������
          if ( unit != null )
            unit.Str = LoadString( "IDS_UNIT" );
        }
        // ��������� ���������
        dim.Update();
      }
    }

  
    //-------------------------------------------------------------------------------
    // �������� � �������������� ������� ���� ����������
    // ---
    void ArcDimensionWork()
    {
	    // �������� ��������� �������� �����������
	    ISymbols2DContainer symbCont = GetSymbols2DContainer();
    	
	    if ( symbCont != null )
	    {
		    // �������� ��������� �������� ���� ����������
		    IArcDimensions arcCol = symbCont.ArcDimensions;
    		
		    if ( arcCol != null )
		    {
			    // �������� ������ ���� ���������� � ���������
			    IArcDimension arcDim = arcCol.Add();
    			
			    if ( arcDim != null )
			    {
				    // ������� ������ ���� ����������
				    bool create = CreateArcDimension( ref arcDim );
				    // �������� �������� �������
				    int dimRef = arcDim.Reference;
    				
				    if ( create && kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
				    {
					    // �������� ������ �� ��������� �� ���������
					    IArcDimension arcDimension = arcCol.get_ArcDimension( dimRef );
					    // ������������� ������ ���� ����������
					    EditArcDimension( ref arcDimension );
				    }
			    }
		    }
	    }	
    }
    #endregion


    #region �������� ������ � �������
    //-------------------------------------------------------------------------------
    // �������� ��������� ������� � �������
    // ---
    bool CreateBreakLineDim( ref IBreakLineDimension dim )
    {
      bool res = false;

	    if ( dim != null && doc2D != null )
	    {
		    // ��������� ���������� ������� � �������
        ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct( 
                             (int)Kompas6Constants.StructType2DEnum.ko_RequestInfo );
        info.Init();
		    info.commandsString = LoadString( "IDS_COMMAND7" );
		    double x = 0, y = 0;
    		
		    // ������� ������ � ���������
		    if ( doc2D.ksCursorEx(info, ref x, ref y, null, null) != 0 )
		    {
			    // ����� ������ �� ��������� �����������
			    int pObj = doc2D.ksFindObj( x, y, doc2D.ksGetCursorLimit() );
			    // ��������� ���������� �������
			    LineSegParam par = (LineSegParam)kompas.GetParamStruct( 
                             (int)Kompas6Constants.StructType2DEnum.ko_LineSegParam );
          par.Init();
    			
			    // �������� �� ������ ��������
			    if ( doc2D.ksGetObjParam(pObj, par, ldefin2d.ALLPARAM ) == ldefin2d.LINESEG_OBJ )
			    {
				    // ���������� ��������� ������
				    doc2D.ksLightObj( pObj, 1 );
				    // �������� ��������� ������������ ������� �� ���������
				    IDrawingObject lineSeg = (IDrawingObject)kompas.TransferReference( pObj, 
                                      doc2D.reference );
    				
				    if ( lineSeg != null )
				    {
					    // ������� ������
					    dim.BaseObject = lineSeg;
					    info.commandsString = LoadString( "IDS_COMMAND6" );					
    					
					    // ������� ��������� ��������� �����
              if ( doc2D.ksCursorEx(info, ref x, ref y, null, null) != 0 )
					    {
						    // ���������� ������ ����� �������� - ���������� ������� ����� �������
						    dim.X1 = par.x1;
						    dim.Y1 = par.y1;
						    // ���������� ������ ����� �������� - ���������� ������� ����� �������
						    dim.X2 = par.x2;
						    dim.Y2 = par.y2;
						    // ��������� ��������� ����� - ��������� ����������
						    dim.X3 = x;
						    dim.Y3 = y;
					    }
					    // ��������� ���������
					    res = dim.Update();
				    }
				    else
              kompas.ksMessage( LoadString("IDS_NOOBJ") );
			    }
			    else
				    kompas.ksMessage( LoadString("IDS_NOTCREATE5") );
    			
			    // ������ ��������� �������
          doc2D.ksLightObj( pObj, 0 );
		    }
	    }
	    return res;
    }


    //-------------------------------------------------------------------------------
    // �������������� ��������� ������� � �������
    // ---
    void EditBreakLineDim( ref IBreakLineDimension dim )
    {
	    if ( dim != null )
	    {
		    // �������� ��������� ���������� �������
		    IDimensionParams dimPar = (IDimensionParams)dim;

		    if ( dimPar != null )
		    {
			    // ��� �������			
			    dimPar.ArrowType1 = Kompas6Constants.ksArrowEnum.ksLeftNotch;
		    }

		    // �������� ��������� ������ �������
		    IDimensionText dimText = (IDimensionText)dim;

		    if ( dimText != null )
		    {
			    // �������� ����� ����� ��������
			    ITextLine prefix = dimText.Prefix;

			    if ( prefix != null )
				    // �������� �����
				    prefix.Str = "������: ";

			    // �������� ����� ������������ ��������
			    ITextLine nominal = dimText.NominalText;
    			
			    if ( nominal != null )
				    // �������� �����
				    nominal.Str = LoadString( "IDS_NOMINAL" );
		    }
		    // ��������� ���������
		    dim.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // �������� � �������������� ��������� ������� � �������
    // ---
    void BreakLineDimensionWork()
    {
	    // �������� ��������� �������� �����������
	    ISymbols2DContainer symbCont = GetSymbols2DContainer();
    	
	    if ( symbCont != null )
	    {
		    // �������� ��������� �������� �������� � �������
		    IBreakLineDimensions breakCol = symbCont.BreakLineDimensions;
    		
		    if ( breakCol != null )
		    {
			    // �������� �������� ������ � ������� � ���������
			    IBreakLineDimension breakDim = breakCol.Add();
    			
			    if ( breakDim != null )
			    {
				    // ������� �������� ������ � �������
				    bool create = CreateBreakLineDim( ref breakDim );
				    // �������� �������� �������
				    int dimRef = breakDim.Reference;
    				
				    if ( create && kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
				    {
					    // �������� ������ �� ��������� �� ���������
					    IBreakLineDimension breakDimension = breakCol.get_BreakLineDimension( dimRef );
					    // ������������� �������� ������ � �������
					    EditBreakLineDim( ref breakDimension );
				    }
			    }
		    }
	    }	
    }
    #endregion


    #region ���������� ������ � �������
    //-------------------------------------------------------------------------------
    // �������� ����������� ������� � �������
    // ---
    bool CreateBreakRadialDim( ref IBreakRadialDimension dim )
    {
      bool res = false;

	    if ( dim != null )
	    {
		    // ��������� ���������� ������� � �������
        ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct( 
                             (int)Kompas6Constants.StructType2DEnum.ko_RequestInfo );
		    info.Init();
		    info.commandsString = LoadString( "IDS_COMMAND2" );
		    double x = 0, y = 0;
    		
		    // ������� ������ � ���������
        if ( doc2D.ksCursorEx(info, ref x, ref y, null, null) != 0 )
		    {
			    // ����� ������ �� ��������� �����������
          int pObj = doc2D.ksFindObj( x, y, doc2D.ksGetCursorLimit() );
			    // ��������� ���������� ����������
			    CircleParam par = (CircleParam)kompas.GetParamStruct( 
                            (int)Kompas6Constants.StructType2DEnum.ko_CircleParam );
    			
			    // �������� �� ������ �����������
			    if ( doc2D.ksGetObjParam(pObj, par, ldefin2d.ALLPARAM ) == ldefin2d.CIRCLE_OBJ )
			    {
				    // �������� ��������� ������������ ������� �� ���������
				    IDrawingObject circle = (IDrawingObject)kompas.TransferReference( pObj, doc2D.reference );
    				
				    if ( circle != null )
				    {
					    // ������ ������� ������ �������
					    dim.BaseObject = circle;
					    // ������ ���������� ������ �������, ����������� � ������� ��������� ����������
					    dim.Xc = par.xc;
					    dim.Yc = par.yc;
					    // ������ ������ �������, ����������� � �������� ��������� ����������
					    dim.Radius = par.rad;
					    // ���� ������� ��������� �����
					    dim.Angle = 45;
					    // ����� ������
					    dim.BreakLength = 3;
					    // ��������� ��������� �������
					    dim.TextOnLine = Kompas6Constants.ksDimensionTextPosEnum.ksDimTextParallelOnLine;
					    // ��������� ���������
					    res = dim.Update();
				    }
				    else
					    kompas.ksMessage( LoadString("IDS_NOOBJ") );
			    }
			    else
				    kompas.ksMessage( LoadString("IDS_NOTCREATE1") );
		    } 
	    }
      return res;
    }


    //-------------------------------------------------------------------------------
    // ������������� ���������� ������ � �������
    // ---
    void EditBreakRadialDim( ref IBreakRadialDimension dim )
    {
	    if ( dim != null )
	    {
		    // ���� ������� ��������� �����
		    dim.Angle = 90;
		    // ����� ������
		    dim.BreakLength = 1;
		    // ��������� ��������� �������
		    dim.TextOnLine = Kompas6Constants.ksDimensionTextPosEnum.ksDimTextParallelInCut;
    	
		    // �������� ��������� ������ �������
		    IDimensionText dimText = (IDimensionText)dim;

		    if ( dimText != null )
		    {
			    // ������ ������ �������
			    dimText.Sign = 0;
		    }
		    // ��������� ���������
		    dim.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // �������� � �������������� ����������� ������� � �������
    // ---
    void BreakRadialDimensionWork()
    {
	    // �������� ��������� �������� �����������
	    ISymbols2DContainer symbCont = GetSymbols2DContainer();
    	
	    if ( symbCont != null )
	    {
		    // �������� ��������� ���������� �������� � �������
		    IBreakRadialDimensions breakCol = symbCont.BreakRadialDimensions;
    		
		    if ( breakCol != null )
		    {
			    // �������� ���������� ������ � ������� � ���������
			    IBreakRadialDimension breakDim = breakCol.Add();
    			
			    if ( breakDim != null )
			    {
				    // ������� ���������� ������ � �������
				    bool create = CreateBreakRadialDim( ref breakDim );
				    // �������� �������� �������
				    int dimRef = breakDim.Reference;
    				
            if ( create && kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
				    {
					    // �������� ������ �� ��������� �� ���������
					    IBreakRadialDimension breakDimension = breakCol.get_BreakRadialDimension( dimRef );
					    // ������������� ���������� ������ � �������
					    EditBreakRadialDim( ref breakDimension );
				    }
			    }
		    }
	    }		
    }
    #endregion


    #region ������� ������ � �������
    //-------------------------------------------------------------------------------
    // �������� �������� ������� � �������
    // ---
    bool CreateBreakAngleDim( ref IBreakAngleDimension dim )
    {
      bool res = false;

	    if ( dim != null && doc2D != null )
	    {
		    // ��������� ������ ������� ������
        int line1 = doc2D.ksLineSeg( 80, 120, 120, 160, 1 );
		    IDrawingObject baseObj1 = (IDrawingObject)kompas.TransferReference( line1, 
                                   doc2D.reference );
		    // ��������� ������ ������� ������
		    int line2 = doc2D.ksLineSeg( 80, 120, 100, 120, 1 );
		    IDrawingObject baseObj2 = (IDrawingObject)kompas.TransferReference( 
                                   line2, doc2D.reference );
    		
		    if ( baseObj1 != null && baseObj2 != null )
		    {
			    // ������ ������� �������
			    dim.BaseObject1 = baseObj1;
			    dim.BaseObject2 = baseObj2;
			    // ���������� ������
			    dim.Xc = 80;
			    dim.Yc = 120;
			    // ������ ����
			    dim.Radius = 20;
			    // ��������� ����  ��������� ����
			    dim.Angle1 = 45;
			    // �������� ����  ��������� ����
			    dim.Angle2 = 30;
			    // ���������� ����� ������ ������ �������� �����
			    dim.X1 = 120;
			    dim.Y1 = 160;
			    // ���������� ����� ������ ������ �������� �����
			    dim.X2 = 100;
			    dim.Y2 = 160;
			    // ��� ������� - �� ����������� (������) ����
			    dim.DimensionType = Kompas6Constants.ksAngleDimTypeEnum.ksADMinAngle;
			    // ����������� ��������� ���� - ������ ������� �������
			    dim.Direction = true;
			    // ������ ����� ��� ����� �� ����
			    dim.X3 = 165;
			    dim.Y3 = 140;
			    // ��������� ���������
			    res = dim.Update();
		    }	
	    }
      return res;
    }


    //-------------------------------------------------------------------------------
    // �������������� �������� ������� � �������
    // ---
    void EditBreakAngleDim( ref IBreakAngleDimension dim )
    {
	    if ( dim != null )
	    {
		    // �������� �����
		    IDimensionParams dimPar = (IDimensionParams)dim;
    		
		    if ( dimPar != null )
			    // ����������� ����� - ������
			    dimPar.ShelfDirection = Kompas6Constants.ksShelfDirectionEnum.ksLSRight;
    		
		    // ����� ������ �����
		    dim.ShelfX = 167;
		    dim.ShelfY = 177;

		    // �������� ��������� ������� �������
		    IDimensionText dimText = (IDimensionText)dim;

		    if ( dimText != null )
		    {
			    // ������������ ��������� �������
			    dimText.TextAlign = Kompas6Constants.ksDimensionTextAlignEnum.ksDimALowerBoundary;
			    // ����� � ������� �������
			    dimText.Brackets = Kompas6Constants.ksDimensionTextBracketsEnum.ksDimBrackets;
		    }
		    // ��������� ���������
		    dim.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // �������� � �������������� �������� ������� � �������
    // ---
    void BreakAngleDimensionWork()
    {
	    // �������� ��������� �������� �����������
	    ISymbols2DContainer symbCont = GetSymbols2DContainer();
    	
	    if ( symbCont != null )
	    {
		    // �������� ��������� ������� �������� � �������
		    IAngleDimensions breakCol = symbCont.AngleDimensions;
    		
		    if ( breakCol != null )
		    {
			    // �������� ������� ������ � ������� � ���������
			    IBreakAngleDimension breakDim = (IBreakAngleDimension)breakCol.Add( 
                                          Kompas6Constants.DrawingObjectTypeEnum.
                                          ksDrABreakDimension );
    			
			    if ( breakDim != null )
			    {
				    // ������� ������� ������ � �������
				    bool create = CreateBreakAngleDim( ref breakDim );
				    // �������� �������� �������
				    int dimRef = breakDim.Reference;
    				
            if ( create && kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
				    {
					    // �������� ������ �� ��������� �� ���������
					    IBreakAngleDimension breakDimension = (IBreakAngleDimension)breakCol.
                                                     get_AngleDimension( dimRef );
					    // ������������� ������� ������ � �������
					    EditBreakAngleDim( ref breakDimension );
				    }
			    }
		    }
	    }	
    }
    #endregion


    #region ������ ������
    //-------------------------------------------------------------------------------
    // ������� ������ ������
    // ---
    bool CreateHeightDimension( ref IHeightDimension dim )
    {
      bool res = false;

	    if ( dim != null )
	    {
		    // ��������� ���������� ������� � �������
        ksRequestInfo info = (ksRequestInfo)kompas.GetParamStruct( (int)Kompas6Constants.StructType2DEnum.ko_RequestInfo );
		    info.Init();
		    info.commandsString = LoadString( "IDS_COMMAND8" );
		    double x = 0, y = 0;
    		
		    // ������� ����� �������� ������
        if ( doc2D.ksCursorEx(info, ref x, ref y, null, null) != 0 )
		    {
			    info.commandsString = LoadString( "IDS_COMMAND9" );
			    double x1 = 0, y1 = 0;

			    // ������� ����� ����������� ������
          if ( doc2D.ksCursorEx(info, ref x1, ref y1, null, null) != 0 )
			    {
				    info.commandsString = LoadString( "IDS_COMMAND10" );
				    double x2 = 0, y2 = 0;
    				
				    // ������� ��������� ��������� �������
            if ( doc2D.ksCursorEx(info, ref x2, ref y2, null, null) != 0 )
				    {
					    // ����� �������� ������
					    dim.X = x;
					    dim.Y = y;
					    // ����� ����������� ������
					    dim.X1 = x1;
					    dim.Y1 = y1;
					    // ��������� ��������� �������
					    dim.X2 = x2;
					    dim.Y2 = y2;
					    // ��������� ���������
					    res = dim.Update();
				    }
			    }
		    }
	    }
      return res;
    }


    //-------------------------------------------------------------------------------
    // ������������� ������ ������
    // ---
    void EditHeightDimension( ref IHeightDimension dim )
    {
	    if ( dim != null )
	    {
		    // ��� �������
		    dim.DimensionType = Kompas6Constants.ksHeightDimTypeEnum.ksHDTopViewLeader;

		    // �������� ��������� ������ �������
		    IDimensionText dimText = (IDimensionText)dim;

		    if ( dimText != null )
		    {
			    // �������� ������ - �������
			    dimText.Sign = 2;
			    // �������� ����� ������������ ��������
			    ITextLine nominal = dimText.NominalText;
    			
			    if ( nominal != null )
				    // �������� �����
				    nominal.Str = LoadString( "IDS_NOMINAL" );
		    }
		    // ��������� ���������
		    dim.Update();
	    }
    }


    //-------------------------------------------------------------------------------
    // �������� � �������������� ������� ������
    // ---
    void HeightDimensionWork()
    {
	    // �������� ��������� �������� �����������
	    ISymbols2DContainer symbCont = GetSymbols2DContainer();
    	
	    if ( symbCont != null )
	    {
		    // �������� ��������� �������� ������
		    IHeightDimensions heightCol = symbCont.HeightDimensions;
    		
		    if ( heightCol != null )
		    {
			    // �������� ������ ������ � ���������
			    IHeightDimension heightDim = heightCol.Add();
    			
			    if ( heightDim != null )
			    {
				    // ������� ������ ������
				    bool create = CreateHeightDimension( ref heightDim );
				    // �������� �������� �������
				    int dimRef = heightDim.Reference;
    				
            if ( create && kompas.ksYesNo(LoadString("IDS_EDIT")) == 1 )
				    {
					    // �������� ������ �� ��������� �� ���������
					    IHeightDimension heightDimension = heightCol.get_HeightDimension( dimRef );
					    // ������������� ������ ������
					    EditHeightDimension( ref heightDimension );
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
