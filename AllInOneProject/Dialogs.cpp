#include "AllinOneScript_rework.h"

int dlg_ColumnNotFound(Worksheet & wks,
					   string s)
{
	set_active_layer(wks);
	
	string strNewCol;
	
	GETN_BOX( badcol );
	
	GETN_STR(STR, "!�� ������ ������� � ������ " + "\"" + s + "\"" + " � ����� " + wks.GetPage().GetLongName() + "!", "") GETN_INFO
	
	GETN_INTERACTIVE(interactive, "�������� ������� � ������� �������", "[Book1]Sheet1!B")
	
	GETN_OPTION_INTERACTIVE_CONTROL(ICOPT_RESTRICT_TO_ONE_DATA) 
	
	if (0 == GetNBox( badcol, "������!" ))
		return -1;
	
	badcol.GetNode("interactive").GetValue(strNewCol);
	
	strNewCol = get_str_from_str_list(strNewCol, "\"", 1);
	
	Column col = ColIdxByName(wks, strNewCol);
	if(!col)
	{
		MessageBox(GetWindow(), "������. ��������, ��������� ���� ������� ������");
		return -1;
	}
	
	//col.SetLongName(s);
	
	return col.GetIndex();
}

int dlg_ZondProcess(vector <int> & vIsProcess,
					vector <int> & vCoefPila,
					vector <int> & vResistance,
					vector & vS,
					vector <string> & vDiagnosticName,
					vector <string> & vBooksNames,
					vector <int> & vSevBooks,
					int & CompareParam)
{
	if (vIsProcess.GetSize() != vCoefPila.GetSize()
		&& vIsProcess.GetSize() != vResistance.GetSize()
		&& vIsProcess.GetSize() != vS.GetSize()
		&& vIsProcess.GetSize() != vDiagnosticName.GetSize())
		return -1;
	
	int NPar = 4, // ���������� ��������� � ������� � ������
		i = 0,
		j = 0; 
	
	/* �������� ������� */
    GETN_BOX( treeTest );
    
    GETN_NUM(freq, "������� ����, ��", freqP)
    
    GETN_BEGIN_BRANCH(FitOpt, "Fitting Options")
    GETN_MULTI_COLS_BRANCH(NPar, DYNA_MULTI_COLS_COMAPCT) GETN_OPTION_BRANCH(GETNBRANCH_OPEN)
    for (i = 0; i < vIsProcess.GetSize(); ++i) 
    {
		GETN_CHECK(ZP,"��������� ������ c " + "\"" + vDiagnosticName[i] + "\"" , vIsProcess[i])
		GETN_NUM(ZC, "����������� �������� ����", vCoefPila[i]) GETN_EDITOR_SIZE_USER_OPTION("#5")
		GETN_NUM(ZR, "������������� �� �����, ��", vResistance[i]) GETN_EDITOR_SIZE_USER_OPTION("#5")
		GETN_NUM(ZS, "������� �����������, �^2", vS[i]) GETN_EDITOR_SIZE_USER_OPTION("#12")
    }
    GETN_END_BRANCH(FitOpt)
    
    GETN_BEGIN_BRANCH(BEP, "������ ������ � ����� ���������(�� ��������� ���������������)") GETN_CHECKBOX_BRANCH(0)
		GETN_SLIDEREDIT(BT, "����� ������, �", 0.3, "0.0|5.0|50") 
		GETN_SLIDEREDIT(ET, "����� �����, �", 2.5, "0.0|5.0|50") 
    GETN_END_BRANCH(BEP)
    
    GETN_NUM(iter, "����������� �������� ������ ����������-����������", Num_iter)
    
    GETN_SLIDEREDIT(left, "����� �����, ������� ����� �������� �� ������", leftP, "0.01|0.89|88")
    GETN_SLIDEREDIT(right, "����� �����, ������� ����� �������� �� �����", rightP, "0.01|0.89|88")
    
    GETN_SLIDEREDIT(linfit, "����� ������ �� ������ ������� ����� ������� �����������������", linfitP, "0.00|0.9|90")
    
    GETN_RADIO_INDEX_EX(IT, "��� ��������", 2, "���|��|���")
    GETN_RADIO_INDEX_EX(FT, "������� ����", 2, "He|Ar|Ne")
    
    GETN_BEGIN_BRANCH(FiltOpt, "���������� ������") GETN_CHECKBOX_BRANCH(1)
    GETN_OPTION_BRANCH(GETNBRANCH_OPEN)
		GETN_SLIDEREDIT(filt, "����� ����� ����������", 0.1, "0.01|0.99|98")
    GETN_END_BRANCH(FiltOpt)
    
    GETN_BEGIN_BRANCH(SevBooks, "����� ���������� ��������� ����") GETN_CHECKBOX_BRANCH(0)
    GETN_RADIO_INDEX_EX(CompareParam, "�������� ��������", 0, "���|��������� ���������|������ ��� ���������|�����������|����������� ���������")
    for (i = 0; i < vBooksNames.GetSize(); ++i) 
    {
		GETN_CHECK(NodeName, vBooksNames[i], true)
    }
    GETN_END_BRANCH(SevBooks)
    
    if (0 == GetNBox( treeTest, "��������� ������ � �����" ))
    	return -1;
    /* ����� �������� ������� */
    
    /* ���������� ������ �� ������� */
    treeTest.GetNode("freq").GetValue(freqP);
    
    TreeNode tn;
    for (i = 0, tn = treeTest.GetNode("FitOpt").GetNode("ZP"); i < vIsProcess.GetSize(); ++i, tn = tn.NextNode) 
    {
    	tn.GetValue(vIsProcess[i]);
    	
    	tn = tn.NextNode;
    	tn.GetValue(vCoefPila[i]);
    	
    	tn = tn.NextNode;
    	tn.GetValue(vResistance[i]);
    	
    	tn = tn.NextNode;
    	tn.GetValue(vS[i]);
    }
	
	if(treeTest.BEP.Use)
	{
		treeTest.GetNode("BEP").GetNode("BT").GetValue(st_time_end_time[0]);
		treeTest.GetNode("BEP").GetNode("ET").GetValue(st_time_end_time[1]);
		if(st_time_end_time[0] >= st_time_end_time[1])
		{
			MessageBox(GetWindow(), "������. ��������� ����� ������ � ����� ���������. �� ������� �� " + st_time_end_time[0] + " �� " + st_time_end_time[1]);
			return -1;
		}
	}
	
	treeTest.GetNode("iter").GetValue(Num_iter);
	
	treeTest.GetNode("left").GetValue(leftP);
	treeTest.GetNode("right").GetValue(rightP);
	
	treeTest.GetNode("linfit").GetValue(linfitP);
	
	treeTest.GetNode("IT").GetValue(ImpulseType);
	treeTest.GetNode("FT").GetValue(FuelType);
	
	if (treeTest.FiltOpt.Use)
		treeTest.GetNode("FiltOpt").GetNode("filt").GetValue(filtS);
	else 
		filtS = 0;
	
	for (i = 0; i < vIsProcess.GetSize(); ++i) 
		if (vIsProcess[i] == 1) 
			j++;
	if (j == 0) 
		vIsProcess[0] = 1;
	
	if(treeTest.SevBooks.Use)
	{
		treeTest.GetNode("SevBooks").GetNode("CompareParam").GetValue(CompareParam);
		vSevBooks.SetSize(vBooksNames.GetSize());
		for (i = 0, tn = treeTest.GetNode("SevBooks").GetNode("NodeName"); i < vBooksNames.GetSize(); ++i, tn = tn.NextNode)
			tn.GetValue(vSevBooks[i])
	}
	/* ����� ���������� ������ �� ������� */
	
	return 0;
}

int dlg_SetkaProcess(vector <int> & vIsProcess,
					 vector <int> & vCoefPila,
					 vector <int> & vResistance,
					 vector <string> & vDiagnosticName,
					 vector <string> & vBooksNames,
					 vector <int> & vSevBooks,
					 int & CompareParam)
{
	if (vIsProcess.GetSize() != vCoefPila.GetSize()
		&& vIsProcess.GetSize() != vResistance.GetSize()
		&& vIsProcess.GetSize() != vDiagnosticName.GetSize())
		return -1;
	
	int NPar = 3, // ���������� ��������� � ������� � ������
		i = 0,
		j = 0; 
	
	/* �������� ������� */
    GETN_BOX( treeTest );
    
    GETN_NUM(freq, "������� ����, ��", freqP)
    
	GETN_BEGIN_BRANCH(FitOpt, "Fitting Options")
    GETN_MULTI_COLS_BRANCH(NPar, DYNA_MULTI_COLS_COMAPCT) GETN_OPTION_BRANCH(GETNBRANCH_OPEN)
    for (i = 0; i < vIsProcess.GetSize(); ++i) 
    {
		GETN_CHECK(SP,"��������� ������ c " + "\"" + vDiagnosticName[i] + "\"" , vIsProcess[i])
		GETN_NUM(SC, "����������� �������� ����", vCoefPila[i]) GETN_EDITOR_SIZE_USER_OPTION("#5")
		GETN_NUM(SR, "������������� �� ��������, ��", vResistance[i]) GETN_EDITOR_SIZE_USER_OPTION("#5")
    }
	GETN_END_BRANCH(FitOpt)
	
    GETN_BEGIN_BRANCH(BEP, "������ ������ � ����� ���������(�� ��������� ���������������)") GETN_CHECKBOX_BRANCH(0)
		GETN_SLIDEREDIT(BT, "����� ������, �", 0.3, "0.0|5.0|50") 
		GETN_SLIDEREDIT(ET, "����� �����, �", 2.5, "0.0|5.0|50") 
    GETN_END_BRANCH(BEP)
    
    GETN_NUM(iter, "����������� �������� ������ ����������-����������", Num_iter)
    
    GETN_SLIDEREDIT(left, "����� �����, ������� ����� �������� �� ������", leftP, "0.01|0.89|88")
    GETN_SLIDEREDIT(right, "����� �����, ������� ����� �������� �� �����", rightP, "0.01|0.89|88")
    
    GETN_RADIO_INDEX_EX(IT, "��� ��������", 2, "���|��|���")
    GETN_RADIO_INDEX_EX(FT, "������� ����", 2, "He|Ar|Ne")
    
    GETN_SLIDEREDIT(filt, "����� ����� ����������", 0.4, "0.01|0.99|98")
    
	GETN_BEGIN_BRANCH(SevBooks, "����� ���������� ��������� ����") GETN_CHECKBOX_BRANCH(0)
	GETN_RADIO_INDEX_EX(CompareParam, "�������� ��������", 0, "���|�������� � ����|�����������|��������� ����|�������")
    for (i = 0; i < vBooksNames.GetSize(); ++i) 
    {
		GETN_CHECK(NodeName, vBooksNames[i], true)
    }
    GETN_END_BRANCH(SevBooks)
    
    if (0 == GetNBox( treeTest, "��������� ������ ���������" ))
    	return -1;
    /* ����� �������� ������� */
    
	/* ���������� ������ �� ������� */
	treeTest.GetNode("freq").GetValue(freqP);
    
    TreeNode tn;
    for (i = 0, tn = treeTest.GetNode("FitOpt").GetNode("SP"); i < vIsProcess.GetSize(); ++i, tn = tn.NextNode) 
    {
    	tn.GetValue(vIsProcess[i]);
    	
    	tn = tn.NextNode;
    	tn.GetValue(vCoefPila[i]);
    	
    	tn = tn.NextNode;
    	tn.GetValue(vResistance[i]);
    }
	
	if(treeTest.BEP.Use)
	{
		treeTest.GetNode("BEP").GetNode("BT").GetValue(st_time_end_time[0]);
		treeTest.GetNode("BEP").GetNode("ET").GetValue(st_time_end_time[1]);
		if(st_time_end_time[0] >= st_time_end_time[1])
		{
			MessageBox(GetWindow(), "������. ��������� ����� ������ � ����� ���������. �� ������� �� " + st_time_end_time[0] + " �� " + st_time_end_time[1]);
			return -1;
		}
	}
	
	treeTest.GetNode("iter").GetValue(Num_iter);
	
	treeTest.GetNode("left").GetValue(leftP);
	treeTest.GetNode("right").GetValue(rightP);
	
	treeTest.GetNode("IT").GetValue(ImpulseType);
	treeTest.GetNode("FT").GetValue(FuelType);
	
	treeTest.GetNode("filt").GetValue(filtS);
	
	for (i = 0; i < vIsProcess.GetSize(); ++i) 
		if (vIsProcess[i] == 1) 
			j++;
	if (j == 0) 
		vIsProcess[0] = 1;
	
	if(treeTest.SevBooks.Use)
	{
		treeTest.GetNode("SevBooks").GetNode("CompareParam").GetValue(CompareParam);
		vSevBooks.SetSize(vBooksNames.GetSize());
		for (i = 0, tn = treeTest.GetNode("SevBooks").GetNode("NodeName"); i < vBooksNames.GetSize(); ++i, tn = tn.NextNode)
			tn.GetValue(vSevBooks[i])
	}
	/* ����� ���������� ������ �� ������� */
	
	return 0;
}

int dlg_CilinProcess(vector <int> & vIsProcess,
					 vector <int> & vCoefPila,
					 vector <int> & vResistance,
					 vector <string> & vDiagnosticName,
					 vector <string> & vBooksNames,
					 vector <int> & vSevBooks,
					 int & CompareParam)
{
	if (vIsProcess.GetSize() != vCoefPila.GetSize()
		&& vIsProcess.GetSize() != vDiagnosticName.GetSize())
		return -1;
	
	int NPar = 3, // ���������� ��������� � ������� � ���������
		i = 0,
		j = 0; 
	
	/* �������� ������� */
    GETN_BOX( treeTest );
    
    GETN_NUM(freq, "������� ����, ��", freqP)
    
    GETN_BEGIN_BRANCH(FitOpt, "Fitting Options")
    GETN_MULTI_COLS_BRANCH(NPar, DYNA_MULTI_COLS_COMAPCT) GETN_OPTION_BRANCH(GETNBRANCH_OPEN)
    for (i = 0; i < vIsProcess.GetSize(); ++i) 
    {
		GETN_CHECK(CP, "��������� ������ c " + "\"" + vDiagnosticName[i] + "\"" , vIsProcess[i])
		GETN_NUM(CC, "����������� �������� ����", vCoefPila[i]) GETN_EDITOR_SIZE_USER_OPTION("#5")
		GETN_NUM(CR, "������������� �������������, ��", vResistance[i]) GETN_EDITOR_SIZE_USER_OPTION("#5")
    }
	GETN_END_BRANCH(FitOpt)
	
    GETN_BEGIN_BRANCH(BEP, "������ ������ � ����� ���������(�� ��������� ���������������)") GETN_CHECKBOX_BRANCH(0)
		GETN_SLIDEREDIT(BT, "����� ������, �", 0.3, "0.0|5.0|50") 
		GETN_SLIDEREDIT(ET, "����� �����, �", 2.5, "0.0|5.0|50") 
    GETN_END_BRANCH(BEP)
    
    GETN_NUM(iter, "����������� �������� ������ ����������-����������", Num_iter)
    
    GETN_SLIDEREDIT(left, "����� �����, ������� ����� �������� �� ������", 0.1, "0.01|0.89|88")
    GETN_SLIDEREDIT(right, "����� �����, ������� ����� �������� �� �����", 0.1, "0.01|0.89|88")
    
    GETN_RADIO_INDEX_EX(IT, "��� ��������", 2, "���|��|���")
    GETN_RADIO_INDEX_EX(FT, "������� ����", 2, "He|Ar|Ne")
    
    GETN_BEGIN_BRANCH(FiltOpt, "���������� ������") GETN_CHECKBOX_BRANCH(1)
    GETN_OPTION_BRANCH(GETNBRANCH_OPEN)
		GETN_SLIDEREDIT(filt, "����� ����� ����������", 0.1, "0.01|0.99|98")
    GETN_END_BRANCH(FiltOpt)
    
	GETN_BEGIN_BRANCH(SevBooks, "����� ���������� ��������� ����") GETN_CHECKBOX_BRANCH(0)
	GETN_RADIO_INDEX_EX(CompareParam, "�������� ��������", 0, "���|�������� � ����|�����������|��������� ����|�������")
    for (i = 0; i < vBooksNames.GetSize(); ++i) 
    {
		GETN_CHECK(NodeName, vBooksNames[i], true)
    }
    GETN_END_BRANCH(SevBooks)
    
    if (0 == GetNBox( treeTest, "��������� ������ � ��������/�������" ))
    	return -1;
    /* ����� �������� ������� */
    
    /* ���������� ������ �� ������� */
	treeTest.GetNode("freq").GetValue(freqP);
	
	TreeNode tn;
    for (i = 0, tn = treeTest.GetNode("FitOpt").GetNode("CP"); i < vIsProcess.GetSize(); ++i, tn = tn.NextNode) 
    {
    	tn.GetValue(vIsProcess[i]);
    	
    	tn = tn.NextNode;
    	tn.GetValue(vCoefPila[i]);
    	
		tn = tn.NextNode;
    	tn.GetValue(vResistance[i]);
    }
	
	if(treeTest.BEP.Use)
	{
		treeTest.GetNode("BEP").GetNode("BT").GetValue(st_time_end_time[0]);
		treeTest.GetNode("BEP").GetNode("ET").GetValue(st_time_end_time[1]);
		if(st_time_end_time[0] >= st_time_end_time[1])
		{
			MessageBox(GetWindow(), "������. ��������� ����� ������ � ����� ���������. �� ������� �� " + st_time_end_time[0] + " �� " + st_time_end_time[1]);
			return -1;
		}
	}
	
	treeTest.GetNode("iter").GetValue(Num_iter);
	
	treeTest.GetNode("left").GetValue(leftP);
	treeTest.GetNode("right").GetValue(rightP);
	
	treeTest.GetNode("IT").GetValue(ImpulseType);
	treeTest.GetNode("FT").GetValue(FuelType);
	
	if(treeTest.FiltOpt.Use)
		treeTest.GetNode("FiltOpt").GetNode("filt").GetValue(filtS);
	else 
		filtS = 0;
	
	for (i = 0; i < vIsProcess.GetSize(); ++i) 
		if (vIsProcess[i] == 1) 
			j++;
	if (j == 0) 
		vIsProcess[0] = 1;
	
	if(treeTest.SevBooks.Use)
	{
		treeTest.GetNode("SevBooks").GetNode("CompareParam").GetValue(CompareParam);
		vSevBooks.SetSize(vBooksNames.GetSize());
		for (i = 0, tn = treeTest.GetNode("SevBooks").GetNode("NodeName"); i < vBooksNames.GetSize(); ++i, tn = tn.NextNode)
			tn.GetValue(vSevBooks[i])
	}
	/* ����� ���������� ������ �� ������� */
	
	return 0;
}

int dlg_DoubleProbeProcess(vector <int> & vIsProcess,
					       vector <int> & vCoefPila,
					       vector <int> & vResistance,
					       vector & vS,
					       vector <string> & vDiagnosticName,
					       vector <string> & vBooksNames,
					       vector <int> & vSevBooks,
						   int & CompareParam)
{
	if (vIsProcess.GetSize() != vCoefPila.GetSize()
		&& vIsProcess.GetSize() != vResistance.GetSize()
		&& vIsProcess.GetSize() != vS.GetSize()
		&& vIsProcess.GetSize() != vDiagnosticName.GetSize())
		return -1;
	
	int NPar = 4, // ���������� ��������� � ������� � ������
		i = 0,
		j = 0; 
	
	/* �������� ������� */
    GETN_BOX( treeTest );
    
    GETN_NUM(freq, "������� ����, ��", freqP)
    
    GETN_BEGIN_BRANCH(FitOpt, "Fitting Options")
    GETN_MULTI_COLS_BRANCH(NPar, DYNA_MULTI_COLS_COMAPCT) GETN_OPTION_BRANCH(GETNBRANCH_OPEN)
    for (i = 0; i < vIsProcess.GetSize(); ++i) 
    {
		GETN_CHECK(ZP,"��������� ������ c " + "\"" + vDiagnosticName[i] + "\"" , vIsProcess[i])
		GETN_NUM(ZC, "����������� �������� ����", vCoefPila[i]) GETN_EDITOR_SIZE_USER_OPTION("#5")
		GETN_NUM(ZR, "������������� �� �����, ��", vResistance[i]) GETN_EDITOR_SIZE_USER_OPTION("#5")
		GETN_NUM(ZS, "������� �����������, �^2", vS[i]) GETN_EDITOR_SIZE_USER_OPTION("#12")
    }
    GETN_END_BRANCH(FitOpt)
    
    GETN_BEGIN_BRANCH(BEP, "������ ������ � ����� ���������(�� ��������� ���������������)") GETN_CHECKBOX_BRANCH(0)
		GETN_SLIDEREDIT(BT, "����� ������, �", 0.3, "0.0|5.0|50") 
		GETN_SLIDEREDIT(ET, "����� �����, �", 2.5, "0.0|5.0|50") 
    GETN_END_BRANCH(BEP)
    
    GETN_NUM(iter, "����������� �������� ������ ����������-����������", Num_iter)
    
    GETN_SLIDEREDIT(left, "����� �����, ������� ����� �������� �� ������", leftP, "0.01|0.89|88")
    GETN_SLIDEREDIT(right, "����� �����, ������� ����� �������� �� �����", rightP, "0.01|0.89|88")
    
    GETN_RADIO_INDEX_EX(IT, "��� ��������", 2, "���|��|���")
    GETN_RADIO_INDEX_EX(FT, "������� ����", 2, "He|Ar|Ne")
    
    GETN_BEGIN_BRANCH(FiltOpt, "���������� ������") GETN_CHECKBOX_BRANCH(1)
    GETN_OPTION_BRANCH(GETNBRANCH_OPEN)
		GETN_SLIDEREDIT(filt, "����� ����� ����������", 0.1, "0.01|0.99|98")
    GETN_END_BRANCH(FiltOpt)
    
    GETN_BEGIN_BRANCH(SevBooks, "����� ���������� ��������� ����") GETN_CHECKBOX_BRANCH(0)
    GETN_RADIO_INDEX_EX(CompareParam, "�������� ��������", 0, "���|����� �������� ����|������ ��� ���������|�����������|����������� ���������")
    for (i = 0; i < vBooksNames.GetSize(); ++i) 
    {
		GETN_CHECK(NodeName, vBooksNames[i], true)
    }
    GETN_END_BRANCH(SevBooks)
    
    if (0 == GetNBox( treeTest, "��������� ������ � �������� �����" ))
    	return -1;
    /* ����� �������� ������� */
    
    /* ���������� ������ �� ������� */
    treeTest.GetNode("freq").GetValue(freqP);
    
    TreeNode tn;
    for (i = 0, tn = treeTest.GetNode("FitOpt").GetNode("ZP"); i < vIsProcess.GetSize(); ++i, tn = tn.NextNode) 
    {
    	tn.GetValue(vIsProcess[i]);
    	
    	tn = tn.NextNode;
    	tn.GetValue(vCoefPila[i]);
    	
    	tn = tn.NextNode;
    	tn.GetValue(vResistance[i]);
    	
    	tn = tn.NextNode;
    	tn.GetValue(vS[i]);
    }
	
	if(treeTest.BEP.Use)
	{
		treeTest.GetNode("BEP").GetNode("BT").GetValue(st_time_end_time[0]);
		treeTest.GetNode("BEP").GetNode("ET").GetValue(st_time_end_time[1]);
		if(st_time_end_time[0] >= st_time_end_time[1])
		{
			MessageBox(GetWindow(), "������. ��������� ����� ������ � ����� ���������. �� ������� �� " + st_time_end_time[0] + " �� " + st_time_end_time[1]);
			return -1;
		}
	}
	
	treeTest.GetNode("iter").GetValue(Num_iter);
	
	treeTest.GetNode("left").GetValue(leftP);
	treeTest.GetNode("right").GetValue(rightP);
	
	treeTest.GetNode("IT").GetValue(ImpulseType);
	treeTest.GetNode("FT").GetValue(FuelType);
	
	if (treeTest.FiltOpt.Use)
		treeTest.GetNode("FiltOpt").GetNode("filt").GetValue(filtS);
	else 
		filtS = 0;
	
	for (i = 0; i < vIsProcess.GetSize(); ++i) 
		if (vIsProcess[i] == 1) 
			j++;
	if (j == 0) 
		vIsProcess[0] = 1;
	
	if(treeTest.SevBooks.Use)
	{
		treeTest.GetNode("SevBooks").GetNode("CompareParam").GetValue(CompareParam);
		vSevBooks.SetSize(vBooksNames.GetSize());
		for (i = 0, tn = treeTest.GetNode("SevBooks").GetNode("NodeName"); i < vBooksNames.GetSize(); ++i, tn = tn.NextNode)
			tn.GetValue(vSevBooks[i])
	}
	/* ����� ���������� ������ �� ������� */
	
	return 0;
}