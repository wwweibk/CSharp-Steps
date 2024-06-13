////////////////////////////////////////////////////////////////////////////////
//
// Object2DEvent - ���������� ������� �������� 2D ���������
//
////////////////////////////////////////////////////////////////////////////////

using Kompas6API5;


using System;
using Kompas6Constants;
using KAPITypes;

namespace Steps.NET
{
	public class Object2DEvent : BaseEvent, ksObject2DNotify
	{
		public Object2DEvent(object obj, object doc, int objType)
			: base(obj, typeof(ksObject2DNotify).GUID, doc, objType, 0) {}

		// kdChangeActive - ������������ ����/���� � �������
		public bool ChangeActive(int viewRef)
		{
			Global.UpdateSlideBox(null);
			return true;
		}


		// koBeginDelete - ������� �������� �������
		public bool BeginDelete(int objRef)
		{
			return true;
		}


		// koDelete - ������ ������
		public bool Delete(int objRef)
		{
			Global.UpdateSlideBox(null);
			return true;
		}
       

		// koBeginMove - ������ �������� �������
		public bool BeginMove(int objRef)
		{
			return true;
		}


		// koMove - ������ ������
		public bool Move(int objRef)
		{
			Global.UpdateSlideBox(null);
			return true;
		}


		// koBeginRotate - ������� �������
		public bool BeginRotate(int objRef)
		{
			return true;
		}


		// koRotate - ������� �������
		public bool Rotate(int objRef)
		{
			Global.UpdateSlideBox(null);
			return true;
		}


		// koBeginScale - �������������� �������
		public bool BeginScale(int objRef)
		{
			return true;
		}


		// koScale - �������������� �������
		public bool scale(int objRef)
		{
			Global.UpdateSlideBox(null);
			return true;
		}


		// koBeginTransform - ������������� �������
		public bool BeginTransform(int objRef)
		{
			return true;
		}


		// koTransform - ������������� �������
		public bool Transform(int objRef)
		{
			Global.UpdateSlideBox(null);
			return true;
		}


		// koBeginCopy - ����������� �������
		public bool BeginCopy(int objRef)
		{
			return true;
		}


		// koCopy - ����������� �������
		public bool copy(int objRef)
		{
			Global.UpdateSlideBox(null);
			return true;
		}


		// koBeginSymmetry - ��������� �������
		public bool BeginSymmetry(int objRef)
		{
			return true;
		}


		// koSymmetry - ��������� �������
		public bool Symmetry(int objRef)
		{
			Global.UpdateSlideBox(null);
			return true;
		}


		// koBeginProcess - ������ ��������������\�������� �������
		public bool BeginProcess(int pType, int objRef)
		{
			return true;
		}
     

		// koEndProcess - ����� ��������������\�������� �������
		public bool EndProcess(int pType)
		{
			Global.UpdateSlideBox(null);
			return true;
		}


		// koCreate - �������� ��������
		public bool CreateObject(int objRef)
		{
			Global.UpdateSlideBox(null);
			return true;
		}

    
		// koUpdateObject - �������������� �������
		public bool UpdateObject(int objRef)
		{
			Global.UpdateSlideBox(null);
			return true;
		}

    // koBeginDestroyObject - ��������� ������
    public bool BeginDestroyObject(int objRef)
    {
      Global.UpdateSlideBox(null);
      return true;
    }

    
    // koDestroyObject - ��������� ������
    public bool DestroyObject(int objRef)
    {
      Global.UpdateSlideBox(null);
      return true;
    }

    public bool BeginPropertyChanged(int objRef)
    {
      return true;
    }

    public bool PropertyChanged(int objRef)
    {
      return true;
    }

		// ������� ���������� ������� ���������
		public static BaseEvent NewObj2DEvent(object doc, int docType, int objType) 
		{
			Object2DEvent res = null;
			if (doc != null) 
			{
				if (!BaseEvent.FindEvent(typeof(Object2DEvent), doc, objType)) 
				{
					ksDocument2D doc2d = (ksDocument2D)doc;
					if (doc != null)
					{
						DocumentEvent.NewDocumentEvent(doc);	// ����� ��� �������� ��������� ����������
						res = new Object2DEvent(doc2d.GetObject2DNotify(objType), doc, objType);
						res.Advise();
					}
				}
			}
			return res;
		}
	}
}
