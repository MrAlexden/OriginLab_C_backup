#include "AllinOneScript_rework.h"

int ch_dsc_p = 50000, 
    ch_dsc_t = 200000, 
    freqP = 50,  
    Num_iter = 50, 
    coefPila = 30, 
    ch_dsc_ntrfrmtr = 2000000, 
    ImpulseType = 0,
    FuelType = 0,
    resistance = 1,
    Number_of_rezult_plasma_parametrs = 5;
double leftP = 0.05, 
       rightP = 0.05,
       razbros_po_pile = 0.25, 
       linfitP = 0.5, 
       st_time_end_time[2] = {-1, -1},
       filtS = 0.1,
       S = 1;	
string s_pz = "���� �����_���", 
	   s_pz2 = "����_����2_���",
   	   s_tz = "��������", 
   	   s_tz2 = "��� ����� 2", 
   	   s_ps = "���� �����_���", 
   	   s_tst = "��� �������� �����", 
   	   s_tssh = "��� �������� ������",
   	   s_tsk = "��� �������� �������",
   	   svch = "���", 
   	   vch = "��", 
   	   icr = "���", 
   	   s_pc = "���� �������", 
   	   s_tc = "�����������������", 
   	   s_ntrfrmtr = "�������������",
   	   s_pm = "�������������", 
	   s_tm = "������������",
	   s_pdz = "���� �����_���"/*"���� �������� �����"*/,
	   s_tdz = "��������"/*"��� �������� �����"*/,
	   s_currp;
  
void zond()
{
	Worksheet wks = Project.ActiveLayer(), wks_slave;
    if (!wks) 
    {
    	MessageBox(GetWindow(), "wks!");
    	return;
    }
    
    int name, CompareParam;
    is_str_numeric_integer(wks.GetPage().GetLongName().Mid(0, 5), &name);
	if (wks.GetPage().GetLongName().Match("����*") 
	   || wks.GetPage().GetLongName().Match("�����*") 
	   || wks.GetPage().GetLongName().Match("�������*") 
	   || wks.GetPage().GetLongName().Match("������*")
	   || wks.GetPage().GetLongName().Match("�����������*")
	   || name < 10000)
	{
		MessageBox(GetWindow(), "�� ��� ����! �������� ���� � ������������ ������������");
		return;
	}
	
	vector <int> vIsProcess = {1, 0}, vSevBooks(0);
	
	vector <int> vCoefPila(vIsProcess.GetSize());
	vCoefPila[0] = 30;
	vCoefPila[1] = 30;
	
	vector <int> vResistance(vIsProcess.GetSize());
	vResistance[0] = 100;
	vResistance[1] = 40;
	
	vector vS(2);
	//vS[0] = 3.141592 * 0.0005 * 0.005 + 3.141592 * 0.0005 * 0.0005 / 4; // ��� ������� ����������� ��������
    vS[0] = 0.0005 * 0.005; // ������ �������� ��� ����������� ������
	//vS[1] = 3.141592*0.002*0.003 + 3.141592*0.002*0.002/4; // ��� ������� ����������� ��������
	vS[1] = 0.002 * 0.003; // ������ �������� ��� ����������� ������
	
	vector <string> vColNamePila(vIsProcess.GetSize());
	vColNamePila[0] = s_pz;
	vColNamePila[1] = s_pz2;
	
	vector <string> vColNameSignal(vIsProcess.GetSize());
	vColNameSignal[0] = s_tz;
	vColNameSignal[1] = s_tz2;
	
	vector <string> vDiagnosticName(vIsProcess.GetSize());
	vDiagnosticName[0] = "����_1";
	vDiagnosticName[1] = "����_2";
	
	vector vecErrs;
	vector <string> vBooksNames,
					vErrsBooksNames,
					vResultBooksNames;
	foreach(WorksheetPage pg in Project.WorksheetPages)
	{
		int BeginVal;
		is_str_numeric_integer(get_str_from_str_list(pg.GetLongName(), "_", 0), &BeginVal);
		if (BeginVal > 10000) 
			vBooksNames.Add(pg.GetLongName());
	}
	if (vBooksNames.GetSize() == 0)
	{
		MessageBox(GetWindow(), "�� ������� ���� � ������������ ������������!");
		return;
	}
	
	if (dlg_ZondProcess(vIsProcess,
						vCoefPila,
						vResistance,
						vS,
						vDiagnosticName,
						vBooksNames,
						vSevBooks,
						CompareParam) < 0)
		return;

	int Error = 0;
	
	if (vSevBooks.GetSize() <= 1)
	{
		vSevBooks.SetSize(1);
		vSevBooks[0] = 1;
		vBooksNames.SetSize(1);
		vBooksNames[0] = wks.GetPage().GetLongName();
	}
	
	for (int j = 0; j < vBooksNames.GetSize(); ++j)
	{
		if (vSevBooks[j] == 0)
			continue;
		
		wks = Project.FindPage(vBooksNames[j]).Layers(0);
		
		for (int i = 0; i < vIsProcess.GetSize(); ++i)
		{
			if (vIsProcess[i] == 0)
				continue;
			
			coefPila = vCoefPila[i];
			resistance = vResistance[i];
			S = vS[i];
			
			Column colPila = ColIdxByName(wks, vColNamePila[i]),
				   colSignal = ColIdxByName(wks, vColNameSignal[i]);
			
			if(!colPila)
			{
				Error = dlg_ColumnNotFound(wks,
										   vColNamePila[i]);
				if (Error < 0)
					return;
				colPila = wks.Columns(Error);
			}
			if(!colSignal)
			{
				Error = dlg_ColumnNotFound(wks,
										   vColNameSignal[i]);
				if (Error < 0)
					return;
				colSignal = wks.Columns(Error);
			}
			
			Error = data_process(0,
								 vDiagnosticName[i],
								 wks, 
								 colPila, 
								 colSignal);
			
			if(Error == 404) return;
			if(Error < 0) 
			{
				vecErrs.Add(Error);
				if(Error != 333)
					vErrsBooksNames.Add("��������� " + vDiagnosticName[i] + " " + wks.GetPage().GetLongName() + ": ");
				else 
					vErrsBooksNames.Add("���� " + vDiagnosticName[i] + " " + wks.GetPage().GetLongName() + ": ");
			}
			Error = 0;
		}
		
		vResultBooksNames.Add(s_currp);
	}
	
	if(vecErrs.GetSize() != 0)
	{
		string str = "���� ������ � ��������� ������: \n";
		
		for(int i = 0; i < vecErrs.GetSize(); ++i) 
			str = str + vErrsBooksNames[i] + ERR_GetErrorDescription(vecErrs[i]) + "\n";
		
		MessageBox(GetWindow(), str);
	}
	
	/* ������� ����� � ���������� ���������� ������ ���������, ���� ���� ������ ������� */
	if (CompareParam > 0)
		compare_books(CompareParam, vResultBooksNames);
	
	////////////////// ��������� ��� ����� � ����������� ��������� ����� �� �� ���������� � �������� //////////////////
	st_time_end_time[0] = -1;
	st_time_end_time[1] = -1;
	////////////////// ��������� ��� ����� � ����������� ��������� ����� �� �� ���������� � �������� //////////////////
}

void setoch()
{
	Worksheet wks = Project.ActiveLayer(), wks_slave;
    if (!wks) 
    {
    	MessageBox(GetWindow(), "wks!");
    	return;
    }
    
	int name, CompareParam;
    is_str_numeric_integer(wks.GetPage().GetLongName().Mid(0, 5), &name);
	if (wks.GetPage().GetLongName().Match("����*") 
	   || wks.GetPage().GetLongName().Match("�����*") 
	   || wks.GetPage().GetLongName().Match("�������*") 
	   || wks.GetPage().GetLongName().Match("������*")
	   || wks.GetPage().GetLongName().Match("�����������*")
	   || name < 10000)
	{
		MessageBox(GetWindow(), "�� ��� ����! �������� ���� � ������������ ������������");
		return;
	}
	
	vector <int> vIsProcess = {1, 0, 0}, vSevBooks(0);
	
	vector <int> vCoefPila(vIsProcess.GetSize());
	vCoefPila[0] = 60;
	vCoefPila[1] = 60;
	vCoefPila[2] = 60;
	
	vector <int> vResistance(vIsProcess.GetSize());
	vResistance[0] = 5000;
	vResistance[1] = 5000;
	vResistance[2] = 5000;
	
	vector <string> vColNamePila(vIsProcess.GetSize());
	vColNamePila[0] = s_ps;
	vColNamePila[1] = s_ps;
	vColNamePila[2] = s_ps;
	
	vector <string> vColNameSignal(vIsProcess.GetSize());
	vColNameSignal[0] = s_tst;
	vColNameSignal[1] = s_tssh;
	vColNameSignal[2] = s_tsk;
	
	vector <string> vDiagnosticName(vIsProcess.GetSize());
	vDiagnosticName[0] = "�����_�";
	vDiagnosticName[1] = "�����_�";
	vDiagnosticName[2] = "�����_�";
	
	vector vecErrs;
	vector <string> vBooksNames,
					vErrsBooksNames,
					vResultBooksNames;
	foreach(WorksheetPage pg in Project.WorksheetPages)
	{
		int BeginVal;
		is_str_numeric_integer(get_str_from_str_list(pg.GetLongName(), "_", 0), &BeginVal);
		if(BeginVal > 10000) 
			vBooksNames.Add(pg.GetLongName());
	}
	if (vBooksNames.GetSize() == 0)
	{
		MessageBox(GetWindow(), "�� ������� ���� � ������������ ������������!");
		return;
	}
	
	if (dlg_SetkaProcess(vIsProcess,
						 vCoefPila,
						 vResistance,
						 vDiagnosticName,
						 vBooksNames,
						 vSevBooks,
						 CompareParam) < 0)
		return;

	int Error = 0;
	
	if (vSevBooks.GetSize() <= 1)
	{
		vSevBooks.SetSize(1);
		vSevBooks[0] = 1;
		vBooksNames.SetSize(1);
		vBooksNames[0] = wks.GetPage().GetLongName();
	}
	
	for (int j = 0; j < vBooksNames.GetSize(); ++j)
	{
		if (vSevBooks[j] == 0)
			continue;
		
		wks = Project.FindPage(vBooksNames[j]).Layers(0);
		
		for (int i = 0; i < vIsProcess.GetSize(); ++i)
		{
			if (vIsProcess[i] == 0)
				continue;
			
			coefPila = vCoefPila[i];
			resistance = vResistance[i];
			
			Column colPila = ColIdxByName(wks, vColNamePila[i]),
				   colSignal = ColIdxByName(wks, vColNameSignal[i]);
			
			if(!colPila)
			{
				Error = dlg_ColumnNotFound(wks,
										   vColNamePila[i]);
				if (Error < 0)
					return;
				colPila = wks.Columns(Error);
			}
			if(!colSignal)
			{
				Error = dlg_ColumnNotFound(wks,
										   vColNameSignal[i]);
				if (Error < 0)
					return;
				colSignal = wks.Columns(Error);
			}
			
			Error = data_process(1,
								 vDiagnosticName[i],
								 wks, 
								 colPila, 
								 colSignal);
			
			if(Error == 404) return;
			if(Error < 0) 
			{
				vecErrs.Add(Error);
				if(Error != 333)
					vErrsBooksNames.Add("��������� " + vDiagnosticName[i] + " " + wks.GetPage().GetLongName() + ": ");
				else 
					vErrsBooksNames.Add("���� " + vDiagnosticName[i] + " " + wks.GetPage().GetLongName() + ": ");
			}
			Error = 0;
		}
		
		vResultBooksNames.Add(s_currp);
	}
	
	if(vecErrs.GetSize() != 0)
	{
		string str = "���� ������ � ��������� ������: \n";
		
		for(int i = 0; i < vecErrs.GetSize(); ++i) 
			str = str + vErrsBooksNames[i] + ERR_GetErrorDescription(vecErrs[i]) + "\n";
		
		MessageBox(GetWindow(), str);
	}
	
	/* ������� ����� � ���������� ���������� ������ ���������, ���� ���� ������ ������� */
	if (CompareParam > 0)
		compare_books(CompareParam, vResultBooksNames);

	////////////////// ��������� ��� ����� � ����������� ��������� ����� �� �� ���������� � �������� //////////////////
	st_time_end_time[0] = -1;
	st_time_end_time[1] = -1;
	////////////////// ��������� ��� ����� � ����������� ��������� ����� �� �� ���������� � �������� //////////////////
}

void cilinder()
{
	Worksheet wks = Project.ActiveLayer(), wks_slave;
    if (!wks) 
    {
    	MessageBox(GetWindow(), "wks!");
    	return;
    }
    
	int name, CompareParam;
    is_str_numeric_integer(wks.GetPage().GetLongName().Mid(0, 5), &name);
	if (wks.GetPage().GetLongName().Match("����*") 
	   || wks.GetPage().GetLongName().Match("�����*") 
	   || wks.GetPage().GetLongName().Match("�������*") 
	   || wks.GetPage().GetLongName().Match("������*")
	   || wks.GetPage().GetLongName().Match("�����������*")
	   || name < 10000)
	{
		MessageBox(GetWindow(), "�� ��� ����! �������� ���� � ������������ ������������");
		return;
	}
	
	vector <int> vIsProcess = {1, 0}, vSevBooks(0);
	
	vector <int> vCoefPila(vIsProcess.GetSize());
	vCoefPila[0] = 66.66;
	vCoefPila[1] = 500;
	
	vector <int> vResistance(vIsProcess.GetSize());
	vResistance[0] = 1;
	vResistance[1] = 1;
	
	vector <string> vColNamePila(vIsProcess.GetSize());
	vColNamePila[0] = s_pc;
	vColNamePila[1] = s_pm;
	
	vector <string> vColNameSignal(vIsProcess.GetSize());
	vColNameSignal[0] = s_tc;
	vColNameSignal[1] = s_tm;
	
	vector <string> vDiagnosticName(vIsProcess.GetSize());
	vDiagnosticName[0] = "�������";
	vDiagnosticName[1] = "������";
	
	vector vecErrs;
	vector <string> vBooksNames,
					vErrsBooksNames,
					vResultBooksNames;
	foreach(WorksheetPage pg in Project.WorksheetPages)
	{
		int BeginVal;
		is_str_numeric_integer(get_str_from_str_list(pg.GetLongName(), "_", 0), &BeginVal);
		if(BeginVal > 10000) 
			vBooksNames.Add(pg.GetLongName());
	}
	if (vBooksNames.GetSize() == 0)
	{
		MessageBox(GetWindow(), "�� ������� ���� � ������������ ������������!");
		return;
	}
	
	if (dlg_CilinProcess(vIsProcess,
						 vCoefPila,
						 vResistance,
						 vDiagnosticName,
						 vBooksNames,
						 vSevBooks,
						 CompareParam) < 0)
		return;

	int Error = 0;
	
	if (vSevBooks.GetSize() <= 1)
	{
		vSevBooks.SetSize(1);
		vSevBooks[0] = 1;
		vBooksNames.SetSize(1);
		vBooksNames[0] = wks.GetPage().GetLongName();
	}
	
	for (int j = 0; j < vBooksNames.GetSize(); ++j)
	{
		if (vSevBooks[j] == 0)
			continue;
		
		wks = Project.FindPage(vBooksNames[j]).Layers(0);
		
		for (int i = 0; i < vIsProcess.GetSize(); ++i)
		{
			if (vIsProcess[i] == 0)
				continue;
			
			coefPila = vCoefPila[i];
			resistance = vResistance[i];
			
			Column colPila = ColIdxByName(wks, vColNamePila[i]),
				   colSignal = ColIdxByName(wks, vColNameSignal[i]);
			
			if(!colPila)
			{
				Error = dlg_ColumnNotFound(wks,
										   vColNamePila[i]);
				if (Error < 0)
					return;
				colPila = wks.Columns(Error);
			}
			if(!colSignal)
			{
				Error = dlg_ColumnNotFound(wks,
										   vColNameSignal[i]);
				if (Error < 0)
					return;
				colSignal = wks.Columns(Error);
			}
			
			Error = data_process(2,
								 vDiagnosticName[i],
								 wks, 
								 colPila, 
								 colSignal);
			
			if(Error == 404) return;
			if(Error < 0) 
			{
				vecErrs.Add(Error);
				if(Error != 333)
					vErrsBooksNames.Add("��������� " + vDiagnosticName[i] + " " + wks.GetPage().GetLongName() + ": ");
				else 
					vErrsBooksNames.Add("���� " + vDiagnosticName[i] + " " + wks.GetPage().GetLongName() + ": ");
			}
			Error = 0;
		}
		
		vResultBooksNames.Add(s_currp);
	}
	
	if(vecErrs.GetSize() != 0)
	{
		string str = "���� ������ � ��������� ������: \n";
		
		for(int i = 0; i < vecErrs.GetSize(); ++i) 
			str = str + vErrsBooksNames[i] + ERR_GetErrorDescription(vecErrs[i]) + "\n";
		
		MessageBox(GetWindow(), str);
	}
	
	/* ������� ����� � ���������� ���������� ������ ���������, ���� ���� ������ ������� */
	if (CompareParam > 0)
		compare_books(CompareParam, vResultBooksNames);
	
	////////////////// ��������� ��� ����� � ����������� ��������� ����� �� �� ���������� � �������� //////////////////
	st_time_end_time[0] = -1;
	st_time_end_time[1] = -1;
	////////////////// ��������� ��� ����� � ����������� ��������� ����� �� �� ���������� � �������� //////////////////
}

void double_probe()
{
	Worksheet wks = Project.ActiveLayer(), wks_slave;
    if (!wks) 
    {
    	MessageBox(GetWindow(), "wks!");
    	return;
    }
    
    int name, CompareParam;
    is_str_numeric_integer(wks.GetPage().GetLongName().Mid(0, 5), &name);
	if (wks.GetPage().GetLongName().Match("����*") 
	   || wks.GetPage().GetLongName().Match("�����*") 
	   || wks.GetPage().GetLongName().Match("�������*") 
	   || wks.GetPage().GetLongName().Match("������*")
	   || wks.GetPage().GetLongName().Match("�����������*")
	   || name < 10000)
	{
		MessageBox(GetWindow(), "�� ��� ����! �������� ���� � ������������ ������������");
		return;
	}
	
	vector <int> vIsProcess = {1}, vSevBooks(0);
	
	vector <int> vCoefPila(vIsProcess.GetSize());
	vCoefPila[0] = 30;
	
	vector <int> vResistance(vIsProcess.GetSize());
	vResistance[0] = 750;
	
	vector vS(1);
	//vS[0] = 3.141592 * 0.0005 * 0.005 + 3.141592 * 0.0005 * 0.0005 / 4; // ��� ������� ����������� ��������
    vS[0] = 0.0005 * 0.005; // ������ �������� ��� ����������� ������
	
	vector <string> vColNamePila(vIsProcess.GetSize());
	vColNamePila[0] = s_pdz;
	
	vector <string> vColNameSignal(vIsProcess.GetSize());
	vColNameSignal[0] = s_tdz;
	
	vector <string> vDiagnosticName(vIsProcess.GetSize());
	vDiagnosticName[0] = "�����������";
	
	vector vecErrs;
	vector <string> vBooksNames,
					vErrsBooksNames,
					vResultBooksNames;
	foreach(WorksheetPage pg in Project.WorksheetPages)
	{
		int BeginVal;
		is_str_numeric_integer(get_str_from_str_list(pg.GetLongName(), "_", 0), &BeginVal);
		if (BeginVal > 10000) 
			vBooksNames.Add(pg.GetLongName());
	}
	if (vBooksNames.GetSize() == 0)
	{
		MessageBox(GetWindow(), "�� ������� ���� � ������������ ������������!");
		return;
	}
	
	if (dlg_DoubleProbeProcess(vIsProcess,
						       vCoefPila,
						       vResistance,
						       vS,
						       vDiagnosticName,
						       vBooksNames,
						       vSevBooks,
							   CompareParam) < 0)
		return;

	int Error = 0;
	
	if (vSevBooks.GetSize() <= 1)
	{
		vSevBooks.SetSize(1);
		vSevBooks[0] = 1;
		vBooksNames.SetSize(1);
		vBooksNames[0] = wks.GetPage().GetLongName();
	}
	
	for (int j = 0; j < vBooksNames.GetSize(); ++j)
	{
		if (vSevBooks[j] == 0)
			continue;
		
		wks = Project.FindPage(vBooksNames[j]).Layers(0);
		
		for (int i = 0; i < vIsProcess.GetSize(); ++i)
		{
			if (vIsProcess[i] == 0)
				continue;
			
			coefPila = vCoefPila[i];
			resistance = vResistance[i];
			S = vS[i];
			
			Column colPila = ColIdxByName(wks, vColNamePila[i]),
				   colSignal = ColIdxByName(wks, vColNameSignal[i]);
			
			if(!colPila)
			{
				Error = dlg_ColumnNotFound(wks,
										   vColNamePila[i]);
				if (Error < 0)
					return;
				colPila = wks.Columns(Error);
			}
			if(!colSignal)
			{
				Error = dlg_ColumnNotFound(wks,
										   vColNameSignal[i]);
				if (Error < 0)
					return;
				colSignal = wks.Columns(Error);
			}
			
			Error = data_process(3,
								 vDiagnosticName[i],
								 wks, 
								 colPila, 
								 colSignal);
			
			if(Error == 404) return;
			if(Error < 0) 
			{
				vecErrs.Add(Error);
				if(Error != 333)
					vErrsBooksNames.Add("��������� " + vDiagnosticName[i] + " " + wks.GetPage().GetLongName() + ": ");
				else 
					vErrsBooksNames.Add("���� " + vDiagnosticName[i] + " " + wks.GetPage().GetLongName() + ": ");
			}
			Error = 0;
		}
		
		vResultBooksNames.Add(s_currp);
	}
	
	if(vecErrs.GetSize() != 0)
	{
		string str = "���� ������ � ��������� ������: \n";
		
		for(int i = 0; i < vecErrs.GetSize(); ++i) 
			str = str + vErrsBooksNames[i] + ERR_GetErrorDescription(vecErrs[i]) + "\n";
		
		MessageBox(GetWindow(), str);
	}
	
	/* ������� ����� � ���������� ���������� ������ ���������, ���� ���� ������ ������� */
	if (CompareParam > 0)
		compare_books(CompareParam, vResultBooksNames);
	
	////////////////// ��������� ��� ����� � ����������� ��������� ����� �� �� ���������� � �������� //////////////////
	st_time_end_time[0] = -1;
	st_time_end_time[1] = -1;
	////////////////// ��������� ��� ����� � ����������� ��������� ����� �� �� ���������� � �������� //////////////////
}

void graph()
{
	int diagnostics;
	
	Page pg;
	
	Worksheet sheetOrigData = Project.ActiveLayer();
	
    if (!sheetOrigData)
    {
    	GraphLayer graph = Project.ActiveLayer();
    	if (!graph)
    	{
    		MessageBox(GetWindow(), "wks! || grapg!");
			return;
    	}
    	
		pg = graph.GetPage();
		
		if (pg.GetLongName().Match("����*")
			|| pg.GetLongName().Match("�����*")
			|| pg.GetLongName().Match("�������*")
			|| pg.GetLongName().Match("������*")
			|| pg.GetLongName().Match("�����������*"))
		{
			if (pg.GetLongName().Match("����*"))
				diagnostics = 0;
			if (pg.GetLongName().Match("�����*"))
				diagnostics = 1;
			if (pg.GetLongName().Match("�������*")
				|| pg.GetLongName().Match("������*"))
				diagnostics = 2;
			if (pg.GetLongName().Match("�����������*"))
				diagnostics = 3;
		}
		else
		{
			MessageBox(GetWindow(), "!���� ��� !����� ��� !������� ��� !������ ��� !����������� (�� ��� ����)");
			return;
		}
		
		if (diagnostics == 0)
			if (reprocess_zond() < 0)
				return;
		if (diagnostics == 1)
			if (reprocess_setka() < 0)
				return;
		if (diagnostics == 2)
			if (reprocess_cilin() < 0)
				return;
		if (diagnostics == 3)
			if (reprocess_doubleprobe() < 0)
				return;
		
		return;
    }
    
    pg = sheetOrigData.GetPage();
    
	if (pg.GetLongName().Match("����*")
		|| pg.GetLongName().Match("�����*")
		|| pg.GetLongName().Match("�������*")
		|| pg.GetLongName().Match("������*")
		|| pg.GetLongName().Match("�����������*"))
	{
		if (pg.GetLongName().Match("����*"))
			diagnostics = 0;
		if (pg.GetLongName().Match("�����*"))
			diagnostics = 1;
		if (pg.GetLongName().Match("�������*")
			|| pg.GetLongName().Match("������*"))
			diagnostics = 2;
		if (pg.GetLongName().Match("�����������*"))
			diagnostics = 3;
	}
	else
	{
		MessageBox(GetWindow(), "!���� ��� !����� ��� !������� ��� !������ ��� !����������� (�� ��� ����)");
		return;
	}
	
	if (diagnostics == 0)
		graphZond();
	if (diagnostics == 1)
		graphSetoch();
	if (diagnostics == 2)
		graphCilin();
	if (diagnostics == 3)
		graphDoubleProbe();
}

void previous_graph()
{
	Page pg;
	
	int z_s_c;

	GraphLayer graph = Project.ActiveLayer();
    if (!graph)
    {
    	MessageBox(GetWindow(), "Grapg! ����� ������� ������ �������������� �� ���� � ��������");
		return;
    }

	pg = graph.GetPage();

	if (pg.GetLongName().Match("����*")
		|| pg.GetLongName().Match("�����*")
		|| pg.GetLongName().Match("�������*")
		|| pg.GetLongName().Match("������*")
		|| pg.GetLongName().Match("�����������*"))
	{
		if (pg.GetLongName().Match("����*")) z_s_c = 0;
		if (pg.GetLongName().Match("�����*")) z_s_c = 1;
		if (pg.GetLongName().Match("�������*")
			|| pg.GetLongName().Match("������*")) z_s_c = 2;
		if (pg.GetLongName().Match("�����������*")) z_s_c = 3;
	}
	else
	{
		MessageBox(GetWindow(), "!���� ��� !����� ��� !������� ��� !������ ��� !����������� (�� ��� ������)");
		return;
	}
	
	if(swap_graph(z_s_c, -1) == 0) return;
}

void next_graph()
{
	Page pg;
	
	int z_s_c;
	
	GraphLayer graph = Project.ActiveLayer();
    if(!graph)
    {
    	MessageBox(GetWindow(), "Grapg! ����� ������� ������ �������������� �� ���� � ��������");
		return;
    }

	pg = graph.GetPage();

	if (pg.GetLongName().Match("����*")
		|| pg.GetLongName().Match("�����*")
		|| pg.GetLongName().Match("�������*")
		|| pg.GetLongName().Match("������*")
		|| pg.GetLongName().Match("�����������*"))
	{
		if (pg.GetLongName().Match("����*")) z_s_c = 0;
		if (pg.GetLongName().Match("�����*")) z_s_c = 1;
		if (pg.GetLongName().Match("�������*")
			|| pg.GetLongName().Match("������*")) z_s_c = 2;
		if (pg.GetLongName().Match("�����������*")) z_s_c = 3;
	}
	else
	{
		MessageBox(GetWindow(), "!���� ��� !����� ��� !������� ��� !������ ��� !����������� (�� ��� ������)");
		return;
	}
	
	if(swap_graph(z_s_c, 1) == 0) return;
}

void interf_transform()
{
	Worksheet wks = Project.ActiveLayer(), wks_slave;
	if (!wks) 
    {
    	MessageBox(GetWindow(), "wks!");
    	return;
    }
	if (wks.GetPage().GetLongName().Match("����*") || wks.GetPage().GetLongName().Match("�����*")
		|| wks.GetPage().GetLongName().Match("�������*") || wks.GetPage().GetLongName().Match("������*"))
	{
		MessageBox(GetWindow(), "�� ��� ����! �������� ���� � ������������ ������������");
		return;
	}
    vector<string> vBooksNames, vErrsBooksNames;
	foreach (WorksheetPage pg in Project.WorksheetPages)
	{
		int BeginVal;
		is_str_numeric_integer(get_str_from_str_list(pg.GetLongName(), "_", 0), &BeginVal);
		if(BeginVal>10000) vBooksNames.Add(pg.GetLongName());
	}
	
	Column colNTRF = ColIdxByName(wks, s_ntrfrmtr);
	vector v;
	v = colNTRF.GetDataObject();
	if (v[0] > 1E5)
	{
		MessageBox(GetWindow(), "�������� ��� ����������!");
		return;
	}
	
	GETN_BOX( treeTest );
    GETN_NUM(decayT1, "������� ������������� ��������������, ��", ch_dsc_ntrfrmtr)
    GETN_RADIO_INDEX_EX(decayT2, "��������� ������� �� ��������� ������", 0, "��|���")
    
    GETN_BEGIN_BRANCH(Details1, "���������� ������") GETN_CHECKBOX_BRANCH(1)
    GETN_OPTION_BRANCH(GETNBRANCH_OPEN)
    GETN_SLIDEREDIT(decayT14, "����� ����� ����������", 0.0001, "0.0001|0.01|98") 
    GETN_END_BRANCH(Details1)
    
    GETN_BEGIN_BRANCH(Details, "����� ���������� ��������� ����") GETN_CHECKBOX_BRANCH(0)
    for(int i_1 = 0; i_1<vBooksNames.GetSize(); i_1++) 
    {
		GETN_CHECK(NodeName, vBooksNames[i_1], true)
    }
    GETN_END_BRANCH(Details)
    
    if(0==GetNBox( treeTest, "��������� ������ � �����" )) return;
    
    time_t start_time, finish_time;
	time(&start_time);
    
    double filtS;
    int vyr_or_not, UseSameName_or_Not_NTRF, True_or_False, Error = -1;
    string ReportString_NTRF;
    vector vecErrs;
    TreeNode tn1;
 
	treeTest.GetNode("decayT1").GetValue(ch_dsc_ntrfrmtr);
	treeTest.GetNode("decayT2").GetValue(vyr_or_not);
	if(treeTest.Details1.Use)
	{
		treeTest.GetNode("Details1").GetNode("decayT14").GetValue(filtS);
		filtS = filtS/2;
	}
	else filtS = 0;
	
	if(!colNTRF)
	{
		GETN_BOX( treeTest1 );
		GETN_STR(STR, "!�� ������ ������� � ������ " + "\"" + s_ntrfrmtr + "\"" + "!", "") GETN_INFO
		GETN_INTERACTIVE(interactive, "�������� ������� � ������� �������", "[Book1]Sheet1!B")
		GETN_OPTION_INTERACTIVE_CONTROL(ICOPT_RESTRICT_TO_ONE_DATA) 
		if(treeTest.Details.Use)
		{
			GETN_RADIO_INDEX_EX(decayT1, "������������ ������� � ����� �� ������ ��� ������� ��������� ����", 0, "��|���")
		}
		if(0==GetNBox( treeTest1, "������!" )) return;
		
		treeTest1.GetNode("interactive").GetValue(ReportString_NTRF);
		treeTest1.GetNode("decayT1").GetValue(UseSameName_or_Not_NTRF);
		ReportString_NTRF = get_str_from_str_list(ReportString_NTRF, "\"", 1);
		colNTRF = ColIdxByName(wks, ReportString_NTRF);
		if(!colNTRF)
		{
			MessageBox(GetWindow(), "������. ��������, ��������� ���� ������� ������");
			return;
		}
		UseSameName_or_Not_NTRF = UseSameName_or_Not_NTRF + 2;
	}
	
	if(treeTest.Details.Use)
	{
		tn1 = treeTest.GetNode("Details").GetNode("NodeName");
		for(int i_2 = 0; i_2<vBooksNames.GetSize(); i_2++)
		{
			if(i_2==0)tn1 = treeTest.GetNode("Details").GetNode("NodeName");
			else tn1 = tn1.NextNode;
			tn1.GetValue(True_or_False);
			if(True_or_False == 1)
			{
				Error = -1;
				Column colNTRF_slave;
				Page pg1 = Project.FindPage(vBooksNames[i_2]);
				if(pg1.Layers(0) != wks) 
				{
					wks_slave = pg1.Layers(0);
					if(UseSameName_or_Not_NTRF == 2) 
					{
						colNTRF_slave = ColIdxByName(wks_slave, ReportString_NTRF);
						if(!colNTRF_slave)
						{
							Error = 333;
							vecErrs.Add(Error);
							vErrsBooksNames.Add("������������� "+vBooksNames[i_2]+": ");
						}
					}
				}
				else
				{
					wks_slave = wks;
					colNTRF_slave = colNTRF;
				}
				if(Error < 0) 
				{
					Error = interf_prog(colNTRF_slave, ch_dsc_ntrfrmtr, wks_slave, filtS, vyr_or_not);
					if(Error == 404) return;
					if(Error != -1) 
					{
						vecErrs.Add(Error);
						vErrsBooksNames.Add("������������� "+vBooksNames[i_2]+": ");
					}
					Error = -1;
				}
			}
		}
	}
	else 
	{
		Error = interf_prog(colNTRF, ch_dsc_ntrfrmtr, wks, filtS, vyr_or_not);
		if(Error == 404) return;
		if( Error != -1 ) 
		{
			MessageBox(GetWindow(), ERR_GetErrorDescription(Error));
			return;
		}
	}
	if(vecErrs.GetSize()!=0)
	{
		string str = "���� ������ � ��������� ������: \n";
		for(int aboba=0; aboba<vecErrs.GetSize(); aboba++) str = str + vErrsBooksNames[aboba] + ERR_GetErrorDescription(vecErrs[aboba]) + "\n";
		MessageBox(GetWindow(), str);
	}
	
	time(&finish_time);
	printf("Run time all prog %d secs\n", finish_time - start_time);
}