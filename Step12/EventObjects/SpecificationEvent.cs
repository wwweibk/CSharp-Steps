////////////////////////////////////////////////////////////////////////////////
//
// SpecificationEvent  - ���������� ������� �� ������������
//
////////////////////////////////////////////////////////////////////////////////

using Kompas6API5;

using System;
using System.Reflection;
using Kompas6Constants;
using KAPITypes;

using reference = System.Int32;

namespace Steps.NET
{
	public class SpecificationEvent : BaseEvent, ksSpecificationNotify
	{
		public SpecificationEvent(object obj, object doc)
			: base(obj, typeof(ksSpecificationNotify).GUID, doc) {}

		// ssTuningSpcStyleBeginChange - ������ ��������� �������� ������������
		public bool TuningSpcStyleBeginChange(string libName, int numb)
		{ 
			return true; 
		}


		// ssTuningSpcStyleChange - ���������� ������������ ����������
		public bool TuningSpcStyleChange(string libName, int numb, bool isOk)
		{
			Global.UpdateSlideBox(null);
			return true;      
		}


		// ssChangeCurrentSpcDescription - ���������� ������� �������� ������������
		public bool ChangeCurrentSpcDescription(string libName, int numb)
		{ 
			//Global.UpdateSlideBox();
			return true; 
		}


		// ssSpcDescriptionAdd - ���������� �������� ������������
		public bool SpcDescriptionAdd(string libName, int numb)
		{ 
			return true;
		}


		// ssSpcDescriptionRemove - ��������� �������� ������������
		public bool SpcDescriptionRemove(string libName, int numb)
		{ 
			return true;
		}


		// ssSpcDescriptionBeginEdit - ������ �������������� �������� ������������
		public bool SpcDescriptionBeginEdit(string libName, int numb)
		{ 
			return true; 
		}


		// ssSpcDescriptionEdit - ��������������� �������� ������������
		public bool SpcDescriptionEdit(string libName, int numb, bool isOk)
		{
			return true; 
		}


		// ssSynchronizationBegin - ������ �������������
		public bool SynchronizationBegin()
		{ 
			return true;
		}


		// ssSynchronization - ������������� ���������
		public bool Synchronization()
		{ 
			return true;
		}


		// ssBeginCalcPositions - ������  ������� �������
		public bool BeginCalcPositions()
		{ 
			return true;
		}


		// ssCalcPositions - �������� ������ ������� 
		public bool CalcPositions()
		{ 
			Global.UpdateSlideBox(null);
			return true;
		}


		// ssBeginCreateObject - ������ �������� ������� ������������ (�� ������� ������ �������) 
		public bool BeginCreateObject(int typeObj)
		{
			return true;
		}


		// ������� ���������� ������� �������������� ������������
		public static BaseEvent NewSpecificationEvent(object doc)
		{
			SpecificationEvent res = null;
			if (doc != null)
			{
				if (!BaseEvent.FindEvent(typeof(SpecificationEvent), doc))
				{
					object spcObj = doc.GetType().InvokeMember("GetSpecification", BindingFlags.InvokeMethod, null, doc, null);
					ksSpecification spc = (ksSpecification)spcObj;
					if (spc != null)
					{
						DocumentEvent.NewDocumentEvent(doc);	// ����� ��� �������� ��������� ����������
						res = new SpecificationEvent(spc, doc);
						res.Advise();
					}
				}
			}
			return res;

		}

	}
}