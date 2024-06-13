////////////////////////////////////////////////////////////////////////////////
//
// SelectMngEvent  - ���������� ������� �� ��������� �������������� ���������
//
////////////////////////////////////////////////////////////////////////////////

using Kompas6API5;

using System;
using Kompas6Constants;
using KAPITypes;

namespace Steps.NET
{
	public class SelectMngEvent : BaseEvent, ksSelectionMngNotify
	{
		public SelectMngEvent(object obj, object doc)
			: base(obj, typeof(ksSelectionMngNotify).GUID, doc) {}

		// ksmSelect - ������ ������������
		public bool Select(object obj) 
		{ 
			Global.UpdateSlideBox(null);
			return true;
		}


		// ksmUnselect - ������ ���������������
		public bool Unselect(object obj)
		{
			Global.UpdateSlideBox(null);
			return true;
		}


		// ksmUnselectAll - ��� ������� ����������������
		public bool UnselectAll() 
		{
			Global.UpdateSlideBox(null);
			return true;
		}	
	}
}
