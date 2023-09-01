#include "AllinOneScript_rework.h"

int dlg_ColumnNotFound(Worksheet & wks,
					   string s)
{
	set_active_layer(wks);
	
	string strNewCol;
	
	GETN_BOX( badcol );
	
	GETN_STR(STR, "!Не найден столбец с именем " + "\"" + s + "\"" + " в книге " + wks.GetPage().GetLongName() + "!", "") GETN_INFO
	
	GETN_INTERACTIVE(interactive, "Выберите столбец с нужными данными", "[Book1]Sheet1!B")
	
	GETN_OPTION_INTERACTIVE_CONTROL(ICOPT_RESTRICT_TO_ONE_DATA) 
	
	if (0 == GetNBox( badcol, "Ошибка!" ))
		return -1;
	
	badcol.GetNode("interactive").GetValue(strNewCol);
	
	strNewCol = get_str_from_str_list(strNewCol, "\"", 1);
	
	Column col = ColIdxByName(wks, strNewCol);
	if(!col)
	{
		MessageBox(GetWindow(), "Ошибка. Возможно, выбранная вами колонка пустая");
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
	
	int NPar = 4, // Количество столбиков в диалоге с зондом
		i = 0,
		j = 0; 
	
	/* СОЗДАНИЕ ДИАЛОГА */
    GETN_BOX( treeTest );
    
    GETN_NUM(freq, "Частота пилы, Гц", freqP)
    
    GETN_BEGIN_BRANCH(FitOpt, "Fitting Options")
    GETN_MULTI_COLS_BRANCH(NPar, DYNA_MULTI_COLS_COMAPCT) GETN_OPTION_BRANCH(GETNBRANCH_OPEN)
    for (i = 0; i < vIsProcess.GetSize(); ++i) 
    {
		GETN_CHECK(ZP,"Обработка данных c " + "\"" + vDiagnosticName[i] + "\"" , vIsProcess[i])
		GETN_NUM(ZC, "Коэффициент усиления пилы", vCoefPila[i]) GETN_EDITOR_SIZE_USER_OPTION("#5")
		GETN_NUM(ZR, "Сопротивление на зонде, Ом", vResistance[i]) GETN_EDITOR_SIZE_USER_OPTION("#5")
		GETN_NUM(ZS, "Площадь поверхности, м^2", vS[i]) GETN_EDITOR_SIZE_USER_OPTION("#12")
    }
    GETN_END_BRANCH(FitOpt)
    
    GETN_BEGIN_BRANCH(BEP, "Задать начало и конец обработки(по умолчанию автоопределение)") GETN_CHECKBOX_BRANCH(0)
		GETN_SLIDEREDIT(BT, "Время начала, с", 0.3, "0.0|5.0|50") 
		GETN_SLIDEREDIT(ET, "Время конца, с", 2.5, "0.0|5.0|50") 
    GETN_END_BRANCH(BEP)
    
    GETN_NUM(iter, "Колическтво итераций метода Левенберга-Марквардта", Num_iter)
    
    GETN_SLIDEREDIT(left, "Часть точек, которые будут отсечены от начала", leftP, "0.01|0.89|88")
    GETN_SLIDEREDIT(right, "Часть точек, которые будут отсечены от конца", rightP, "0.01|0.89|88")
    
    GETN_SLIDEREDIT(linfit, "Часть данных от начала которые будут линейно аппроксимированны", linfitP, "0.00|0.9|90")
    
    GETN_RADIO_INDEX_EX(IT, "Тип импульса", 2, "СВЧ|ВЧ|ИЦР")
    GETN_RADIO_INDEX_EX(FT, "Рабочее тело", 2, "He|Ar|Ne")
    
    GETN_BEGIN_BRANCH(FiltOpt, "Фильтрация данных") GETN_CHECKBOX_BRANCH(1)
    GETN_OPTION_BRANCH(GETNBRANCH_OPEN)
		GETN_SLIDEREDIT(filt, "Часть точек фильтрации", 0.1, "0.01|0.99|98")
    GETN_END_BRANCH(FiltOpt)
    
    GETN_BEGIN_BRANCH(SevBooks, "Нужно обработать несколько книг") GETN_CHECKBOX_BRANCH(0)
    GETN_RADIO_INDEX_EX(CompareParam, "Сравнить параметр", 0, "Нет|Плавающий потенциал|Ионный ток насыщения|Температура|Электронная плотность")
    for (i = 0; i < vBooksNames.GetSize(); ++i) 
    {
		GETN_CHECK(NodeName, vBooksNames[i], true)
    }
    GETN_END_BRANCH(SevBooks)
    
    if (0 == GetNBox( treeTest, "Обработка данных с зонда" ))
    	return -1;
    /* КОНЕЦ СОЗДАНИЯ ДИАЛОГА */
    
    /* СЧИТЫВАНИЕ ДАННЫХ ИЗ ДИАЛОГА */
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
			MessageBox(GetWindow(), "Ошибка. Исправьте время начала и конца обработки. Вы выбрали от " + st_time_end_time[0] + " до " + st_time_end_time[1]);
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
	/* КОНЕЦ СЧИТЫВАНИЯ ДАННЫХ ИЗ ДИАЛОГА */
	
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
	
	int NPar = 3, // Количество столбиков в диалоге с сеткой
		i = 0,
		j = 0; 
	
	/* СОЗДАНИЕ ДИАЛОГА */
    GETN_BOX( treeTest );
    
    GETN_NUM(freq, "Частота пилы, Гц", freqP)
    
	GETN_BEGIN_BRANCH(FitOpt, "Fitting Options")
    GETN_MULTI_COLS_BRANCH(NPar, DYNA_MULTI_COLS_COMAPCT) GETN_OPTION_BRANCH(GETNBRANCH_OPEN)
    for (i = 0; i < vIsProcess.GetSize(); ++i) 
    {
		GETN_CHECK(SP,"Обработка данных c " + "\"" + vDiagnosticName[i] + "\"" , vIsProcess[i])
		GETN_NUM(SC, "Коэффициент усиления пилы", vCoefPila[i]) GETN_EDITOR_SIZE_USER_OPTION("#5")
		GETN_NUM(SR, "Сопротивление на сеточном, Ом", vResistance[i]) GETN_EDITOR_SIZE_USER_OPTION("#5")
    }
	GETN_END_BRANCH(FitOpt)
	
    GETN_BEGIN_BRANCH(BEP, "Задать начало и конец обработки(по умолчанию автоопределение)") GETN_CHECKBOX_BRANCH(0)
		GETN_SLIDEREDIT(BT, "Время начала, с", 0.3, "0.0|5.0|50") 
		GETN_SLIDEREDIT(ET, "Время конца, с", 2.5, "0.0|5.0|50") 
    GETN_END_BRANCH(BEP)
    
    GETN_NUM(iter, "Колическтво итераций метода Левенберга-Марквардта", Num_iter)
    
    GETN_SLIDEREDIT(left, "Часть точек, которые будут отсечены от начала", leftP, "0.01|0.89|88")
    GETN_SLIDEREDIT(right, "Часть точек, которые будут отсечены от конца", rightP, "0.01|0.89|88")
    
    GETN_RADIO_INDEX_EX(IT, "Тип импульса", 2, "СВЧ|ВЧ|ИЦР")
    GETN_RADIO_INDEX_EX(FT, "Рабочее тело", 2, "He|Ar|Ne")
    
    GETN_SLIDEREDIT(filt, "Часть точек фильтрации", 0.4, "0.01|0.99|98")
    
	GETN_BEGIN_BRANCH(SevBooks, "Нужно обработать несколько книг") GETN_CHECKBOX_BRANCH(0)
	GETN_RADIO_INDEX_EX(CompareParam, "Сравнить параметр", 0, "Нет|Значение в пике|Температура|Потенциал пика|Энергия")
    for (i = 0; i < vBooksNames.GetSize(); ++i) 
    {
		GETN_CHECK(NodeName, vBooksNames[i], true)
    }
    GETN_END_BRANCH(SevBooks)
    
    if (0 == GetNBox( treeTest, "Обработка данных сеточного" ))
    	return -1;
    /* КОНЕЦ СОЗДАНИЯ ДИАЛОГА */
    
	/* СЧИТЫВАНИЕ ДАННЫХ ИЗ ДИАЛОГА */
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
			MessageBox(GetWindow(), "Ошибка. Исправьте время начала и конца обработки. Вы выбрали от " + st_time_end_time[0] + " до " + st_time_end_time[1]);
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
	/* КОНЕЦ СЧИТЫВАНИЯ ДАННЫХ ИЗ ДИАЛОГА */
	
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
	
	int NPar = 3, // Количество столбиков в диалоге с цилиндром
		i = 0,
		j = 0; 
	
	/* СОЗДАНИЕ ДИАЛОГА */
    GETN_BOX( treeTest );
    
    GETN_NUM(freq, "Частота пилы, Гц", freqP)
    
    GETN_BEGIN_BRANCH(FitOpt, "Fitting Options")
    GETN_MULTI_COLS_BRANCH(NPar, DYNA_MULTI_COLS_COMAPCT) GETN_OPTION_BRANCH(GETNBRANCH_OPEN)
    for (i = 0; i < vIsProcess.GetSize(); ++i) 
    {
		GETN_CHECK(CP, "Обработка данных c " + "\"" + vDiagnosticName[i] + "\"" , vIsProcess[i])
		GETN_NUM(CC, "Коэффициент усиления пилы", vCoefPila[i]) GETN_EDITOR_SIZE_USER_OPTION("#5")
		GETN_NUM(CR, "Измерительное сопротивление, Ом", vResistance[i]) GETN_EDITOR_SIZE_USER_OPTION("#5")
    }
	GETN_END_BRANCH(FitOpt)
	
    GETN_BEGIN_BRANCH(BEP, "Задать начало и конец обработки(по умолчанию автоопределение)") GETN_CHECKBOX_BRANCH(0)
		GETN_SLIDEREDIT(BT, "Время начала, с", 0.3, "0.0|5.0|50") 
		GETN_SLIDEREDIT(ET, "Время конца, с", 2.5, "0.0|5.0|50") 
    GETN_END_BRANCH(BEP)
    
    GETN_NUM(iter, "Колическтво итераций метода Левенберга-Марквардта", Num_iter)
    
    GETN_SLIDEREDIT(left, "Часть точек, которые будут отсечены от начала", 0.1, "0.01|0.89|88")
    GETN_SLIDEREDIT(right, "Часть точек, которые будут отсечены от конца", 0.1, "0.01|0.89|88")
    
    GETN_RADIO_INDEX_EX(IT, "Тип импульса", 2, "СВЧ|ВЧ|ИЦР")
    GETN_RADIO_INDEX_EX(FT, "Рабочее тело", 2, "He|Ar|Ne")
    
    GETN_BEGIN_BRANCH(FiltOpt, "Фильтрация данных") GETN_CHECKBOX_BRANCH(1)
    GETN_OPTION_BRANCH(GETNBRANCH_OPEN)
		GETN_SLIDEREDIT(filt, "Часть точек фильтрации", 0.1, "0.01|0.99|98")
    GETN_END_BRANCH(FiltOpt)
    
	GETN_BEGIN_BRANCH(SevBooks, "Нужно обработать несколько книг") GETN_CHECKBOX_BRANCH(0)
	GETN_RADIO_INDEX_EX(CompareParam, "Сравнить параметр", 0, "Нет|Значение в пике|Температура|Потенциал пика|Энергия")
    for (i = 0; i < vBooksNames.GetSize(); ++i) 
    {
		GETN_CHECK(NodeName, vBooksNames[i], true)
    }
    GETN_END_BRANCH(SevBooks)
    
    if (0 == GetNBox( treeTest, "Обработка данных с цилиндра/магнита" ))
    	return -1;
    /* КОНЕЦ СОЗДАНИЯ ДИАЛОГА */
    
    /* СЧИТЫВАНИЕ ДАННЫХ ИЗ ДИАЛОГА */
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
			MessageBox(GetWindow(), "Ошибка. Исправьте время начала и конца обработки. Вы выбрали от " + st_time_end_time[0] + " до " + st_time_end_time[1]);
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
	/* КОНЕЦ СЧИТЫВАНИЯ ДАННЫХ ИЗ ДИАЛОГА */
	
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
	
	int NPar = 4, // Количество столбиков в диалоге с зондом
		i = 0,
		j = 0; 
	
	/* СОЗДАНИЕ ДИАЛОГА */
    GETN_BOX( treeTest );
    
    GETN_NUM(freq, "Частота пилы, Гц", freqP)
    
    GETN_BEGIN_BRANCH(FitOpt, "Fitting Options")
    GETN_MULTI_COLS_BRANCH(NPar, DYNA_MULTI_COLS_COMAPCT) GETN_OPTION_BRANCH(GETNBRANCH_OPEN)
    for (i = 0; i < vIsProcess.GetSize(); ++i) 
    {
		GETN_CHECK(ZP,"Обработка данных c " + "\"" + vDiagnosticName[i] + "\"" , vIsProcess[i])
		GETN_NUM(ZC, "Коэффициент усиления пилы", vCoefPila[i]) GETN_EDITOR_SIZE_USER_OPTION("#5")
		GETN_NUM(ZR, "Сопротивление на зонде, Ом", vResistance[i]) GETN_EDITOR_SIZE_USER_OPTION("#5")
		GETN_NUM(ZS, "Площадь поверхности, м^2", vS[i]) GETN_EDITOR_SIZE_USER_OPTION("#12")
    }
    GETN_END_BRANCH(FitOpt)
    
    GETN_BEGIN_BRANCH(BEP, "Задать начало и конец обработки(по умолчанию автоопределение)") GETN_CHECKBOX_BRANCH(0)
		GETN_SLIDEREDIT(BT, "Время начала, с", 0.3, "0.0|5.0|50") 
		GETN_SLIDEREDIT(ET, "Время конца, с", 2.5, "0.0|5.0|50") 
    GETN_END_BRANCH(BEP)
    
    GETN_NUM(iter, "Колическтво итераций метода Левенберга-Марквардта", Num_iter)
    
    GETN_SLIDEREDIT(left, "Часть точек, которые будут отсечены от начала", leftP, "0.01|0.89|88")
    GETN_SLIDEREDIT(right, "Часть точек, которые будут отсечены от конца", rightP, "0.01|0.89|88")
    
    GETN_RADIO_INDEX_EX(IT, "Тип импульса", 2, "СВЧ|ВЧ|ИЦР")
    GETN_RADIO_INDEX_EX(FT, "Рабочее тело", 2, "He|Ar|Ne")
    
    GETN_BEGIN_BRANCH(FiltOpt, "Фильтрация данных") GETN_CHECKBOX_BRANCH(1)
    GETN_OPTION_BRANCH(GETNBRANCH_OPEN)
		GETN_SLIDEREDIT(filt, "Часть точек фильтрации", 0.1, "0.01|0.99|98")
    GETN_END_BRANCH(FiltOpt)
    
    GETN_BEGIN_BRANCH(SevBooks, "Нужно обработать несколько книг") GETN_CHECKBOX_BRANCH(0)
    GETN_RADIO_INDEX_EX(CompareParam, "Сравнить параметр", 0, "Нет|Точка нулевого тока|Ионный ток насыщения|Температура|Электронная плотность")
    for (i = 0; i < vBooksNames.GetSize(); ++i) 
    {
		GETN_CHECK(NodeName, vBooksNames[i], true)
    }
    GETN_END_BRANCH(SevBooks)
    
    if (0 == GetNBox( treeTest, "Обработка данных с двойного зонда" ))
    	return -1;
    /* КОНЕЦ СОЗДАНИЯ ДИАЛОГА */
    
    /* СЧИТЫВАНИЕ ДАННЫХ ИЗ ДИАЛОГА */
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
			MessageBox(GetWindow(), "Ошибка. Исправьте время начала и конца обработки. Вы выбрали от " + st_time_end_time[0] + " до " + st_time_end_time[1]);
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
	/* КОНЕЦ СЧИТЫВАНИЯ ДАННЫХ ИЗ ДИАЛОГА */
	
	return 0;
}