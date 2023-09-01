#include "AllinOneScript_rework.h"

void swap_i(int & first, int & second)
{
	int firstbuff = first, secondbuff = second;
	
	first = secondbuff;
	second = firstbuff;
}

string ERR_GetErrorDescription(int err)
{
    switch (err)
    {
	case -222: 
		return "Ошибка! Не найдена колонка с током или она пустая";
	case -333: 
		return "Ошибка! Не найдена колонка с пилой или она пустая";
	case -555: 
		return "Ошибка! Пересчет плотности уже применялся";
	case -666: 
		return "Ошибка! Не найдена колонка с интерферометром или она пустая";
    case ERR_BadInputVecs:
        return "Corrupted input vectors";
    case ERR_ZeroInputVals:
        return "Input data error, some values equals 0";
    case ERR_BadCutOff:
        return "Сut-off points must be less than 0.9 in total and more than 0 each";
    case ERR_BadLinFit:
        return "The length of linear fit must be less than 0.9";
    case ERR_BadFactorizing:
        return "Error after Pila|Signal factorizing";
    case ERR_BadNoise:
        return "Error after noise extracting";
    case ERR_BadSegInput:
        return "Input segment's values error while segmend approximating";
    case ERR_TooFewSegs:
        return "Less then 4 segments found, check input arrays";
    case ERR_BadSegsLength:
        return "Error in finding segments length, check input params";
    case ERR_BadLinearPila:
        return "Error in pila linearizing, check cut-off params";
    case ERR_TooManyAttempts:
        return "More than 5 attempts to find signal, check if signal is noise or not";
    case ERR_BadStartEnd:
        return "Error in finding start|end of signal, check if signal is noise or not";
    case ERR_TooManySegs:
        return "Too many segments, check if signal is noise or not";
    case ERR_NoSegs:
        return "No segments found, check if signal is noise or not";
    case ERR_BadDiagNum:
        return "Diagnostics number must be > 0 and < 2";
    case ERR_IdxOutOfRange:
        return "Index is out of range. Programm continued running";
    case ERR_Exception:
        return "An exception been occured while running, script did not stoped working";
    case ERR_BadStEndTime:
        return "Error!: start time must be less then end time and total time, more than 0\n\
end time must be less then total time, more then 0";
	case ERR_BufferExtension:
        return "Buffer extension, bad impulse";
    default:
        return "No Error";
    }
    
    return "No Error";
}

string FindPageNameNumber(string newPgName)
{
	Project proj;
	GraphPage gp;
	GraphLayer graph;
	
	Worksheet wks = Project.ActiveLayer();
	if(!wks) 
	{
		graph = Project.ActiveLayer();
		graph.GetPage().GetParent(proj);
	}
	else wks.GetPage().GetParent(proj);
	
	int i=1,k;
	while(i>0)
	{
		k=i;
		foreach(PageBase pg in Project.Pages)
		{
			if(pg.GetLongName().CompareNoCase(newPgName + " №" + i) == 0) 
			{
				i++;
				break;
			}
		}
		if(k==i)break;
	}
    return newPgName + " №" + i;
}

Worksheet & WksIdxByName(Page & pg, string strLabel)
{
	foreach(Layer wks in pg.Layers)
		if (wks.GetName().CompareNoCase(strLabel) == 0)
			return wks;
}

Column & ColIdxByName(Worksheet & wks, string strLabel)
{
	foreach(Column col in wks.Columns)
		if (col.GetLongName().CompareNoCase(strLabel) == 0 && col.GetUpperBound() > 0) 
			return col;	
}

Column & ColIdxByUnit(Worksheet & wks, string strLabel)
{
	foreach(Column col in wks.Columns)
		if (get_str_from_str_list(col.GetUnits(), " ", 0).CompareNoCase(strLabel) == 0)
			return col;
}

int Get_OriginData_wks(Page & pgProcessData, Page & pgOrigData, Page & pgLegend)
{
	Worksheet wksProcessData,
			  wksLegend,
			  wksOrigData;
	
	GraphLayer graphActive = Project.ActiveLayer(), graph = graphActive.GetPage().Layers(0);
	if(!graph) 
	{
		MessageBox(GetWindow(), "!graph");
		return -1;
	}
	
	vector<uint> vUIDs;
    if(graph.GetConnectedObjects(vUIDs)<=1) 
    {
    	MessageBox(GetWindow(), "Не найден родительский объект графика или книга легенды");
		return -1;
    }
    
    wksProcessData = Project.GetObject(vUIDs[0]);
    if(!wksProcessData) 
    {
    	MessageBox(GetWindow(), "Не найден родительский объект графика");
		return -1;
    }
    pgProcessData = wksProcessData.GetPage();
    
    wksLegend = Project.GetObject(vUIDs[1]);
    if(!wksLegend) 
    {
    	MessageBox(GetWindow(), "Не найдена книга легенды");
		return -1;
    }
    pgLegend = wksLegend.GetPage();
    
    vUIDs.RemoveAll();
    if(wksProcessData.GetConnectedObjects(vUIDs)<=0)
    {
    	MessageBox(GetWindow(), "Не найдена книга с оригинальными данными");
		return -1;
    }
    
    for(int ii=0; ii<vUIDs.GetSize(); ii++)
    {
        wksOrigData = Project.GetObject(vUIDs[ii]);
        if(wksOrigData)
        {
        	pgOrigData = wksOrigData.GetPage();
        	return 0;
        }
    }
    
	if(!wksOrigData)
	{
		MessageBox(GetWindow(), "Недействительна книга с оригинальными данными");
		return -1;
	}
	
    return 0;
}

Page & create_book(int diagnostics,
				   string diagnosticName,
				   int DIM1,
				   int DIM2,
				   vector &vTime,
				   vector <int> &vStartSegIndxs,
				   string pgOrigName,
				   string ColNamePila,
				   string ColNameSignal)				    
{
	/* create book */
	Worksheet sheetOrigData, sheetFitData, sheetParamsData, sheetFiltData, sheetDiffData;
	sheetOrigData.Create("Origin");
	Page pg = sheetOrigData.GetPage();
	
	string newPgName, strUnits, strImpulseType, strFuelType, strFilt, strCutOffs;
	switch (ImpulseType)
	{
		case 0:
			strImpulseType = svch;
			break;
		case 1:
			strImpulseType = vch;
			break;
		case 2:
			strImpulseType = icr;
			break;
	}
	switch (FuelType)
	{
		case 0:
			strFuelType = "He";
			break;
		case 1:
			strFuelType = "Ar";
			break;
		case 2:
			strFuelType = "Ne";
			break;
	}
	
	strUnits.Format("%s %s %i", strImpulseType, strFuelType, resistance);
	strCutOffs.Format("B %.2f E %.2f LF %.2f", leftP, rightP, linfitP);
	
	/* create 1st sheet */
	sheetOrigData = pg.Layers(pg.Layers.Count() - 1);
	while (sheetOrigData.DeleteCol(0));
	sheetOrigData.SetSize(DIM2, DIM1 + 1);
	sheetOrigData.SetName("OrigData");
	sheetOrigData.Columns(0).SetComments("0");
	sheetOrigData.Columns(0).SetType(OKDATAOBJ_DESIGNATION_X);
	sheetOrigData.Columns(0).SetUnits(strUnits);
	sheetOrigData.Columns(0).SetLongName("pila");
	
	/* create 2nd sheet optionally */
	if (filtS != 0)
	{
		pg.AddLayer("Filtration");
		sheetFiltData = pg.Layers(pg.Layers.Count() - 1);
		while (sheetFiltData.DeleteCol(0));
		sheetFiltData.SetSize(DIM2, DIM1 + 1);
		sheetFiltData.Columns(0).SetComments("0");
		sheetFiltData.Columns(0).SetType(OKDATAOBJ_DESIGNATION_X);
		sheetFiltData.Columns(0).SetUnits(strUnits);
		sheetFiltData.Columns(0).SetLongName("pila");
		
		strFilt.Format("%.2f", filtS);
	}
	
	/* 3rd sheet only for setka */
	if (diagnostics == 1)
	{
		pg.AddLayer("Diff");
		sheetDiffData = pg.Layers(pg.Layers.Count() - 1);
		while (sheetDiffData.DeleteCol(0));
		sheetDiffData.SetSize(DIM2, DIM1 + 1);
		sheetDiffData.Columns(0).SetComments("0");
		sheetDiffData.Columns(0).SetType(OKDATAOBJ_DESIGNATION_X);
		sheetDiffData.Columns(0).SetUnits(strUnits);
		sheetDiffData.Columns(0).SetLongName("pila");
	}
	
	/* create 4th sheet */
	pg.AddLayer("Fit");
	sheetFitData = pg.Layers(pg.Layers.Count() - 1);
	while (sheetFitData.DeleteCol(0));
	sheetFitData.SetSize(DIM2, DIM1 + 1);
	sheetFitData.Columns(0).SetComments("0");
	sheetFitData.Columns(0).SetType(OKDATAOBJ_DESIGNATION_X);
	sheetFitData.Columns(0).SetUnits(strUnits);
	sheetFitData.Columns(0).SetLongName("pila");
	
	/* create 5th sheet */
	pg.AddLayer("Parameters");
	sheetParamsData = pg.Layers(pg.Layers.Count() - 1);
	while (sheetParamsData.DeleteCol(0));
	sheetParamsData.SetSize(DIM1, 5);
	sheetParamsData.Columns(0).SetComments(ColNamePila + "|" + ColNameSignal);
	sheetParamsData.Columns(0).SetType(OKDATAOBJ_DESIGNATION_X);
	sheetParamsData.Columns(0).SetUnits(strUnits + " " + S);
	sheetParamsData.Columns(0).SetLongName("Время");
	
	switch (diagnostics)
	{
		case 0:
			newPgName.Format(diagnosticName + " %s %s %s", strImpulseType, strFuelType, pgOrigName);
			pg.SetLongName(FindPageNameNumber(newPgName));
			
			sheetParamsData.Columns(1).SetLongName("Плавающий потенциал " + pgOrigName);
			sheetParamsData.Columns(1).SetUnits("В");
			sheetParamsData.Columns(2).SetLongName("Ионный ток насыщения " + pgOrigName);
			sheetParamsData.Columns(2).SetUnits("А");
			sheetParamsData.Columns(3).SetLongName("Температура " + pgOrigName);
			sheetParamsData.Columns(3).SetUnits("эВ");
			sheetParamsData.Columns(4).SetLongName("Электронная плотность " + pgOrigName);
			sheetParamsData.Columns(4).SetUnits("см^-3");
			break;
		case 1:
			newPgName.Format(diagnosticName + " %s %s %s", strImpulseType, strFuelType, pgOrigName);
			pg.SetLongName(FindPageNameNumber(newPgName));
			
			sheetParamsData.Columns(1).SetLongName("Максимальный ток " + pgOrigName);
			sheetParamsData.Columns(1).SetUnits("А");
			sheetParamsData.Columns(2).SetLongName("Температура " + pgOrigName);
			sheetParamsData.Columns(2).SetUnits("эВ");
			sheetParamsData.Columns(3).SetLongName("Потенциал пика " + pgOrigName);
			sheetParamsData.Columns(3).SetUnits("В");
			sheetParamsData.Columns(4).SetLongName("Энергия " + pgOrigName);
			sheetParamsData.Columns(4).SetUnits("эВ");
			break;
		case 2:
			newPgName.Format(diagnosticName + " %s %s %s", strImpulseType, strFuelType, pgOrigName);
			pg.SetLongName(FindPageNameNumber(newPgName));
			
			sheetParamsData.Columns(1).SetLongName("Максимальный ток " + pgOrigName);
			sheetParamsData.Columns(1).SetUnits("А");
			sheetParamsData.Columns(2).SetLongName("Температура " + pgOrigName);
			sheetParamsData.Columns(2).SetUnits("эВ");
			sheetParamsData.Columns(3).SetLongName("Потенциал пика " + pgOrigName);
			sheetParamsData.Columns(3).SetUnits("В");
			sheetParamsData.Columns(4).SetLongName("Энергия " + pgOrigName);
			sheetParamsData.Columns(4).SetUnits("эВ");
			break;
		case 3:
			newPgName.Format(diagnosticName + " %s %s %s", strImpulseType, strFuelType, pgOrigName);
			pg.SetLongName(FindPageNameNumber(newPgName));
			
			sheetParamsData.Columns(1).SetLongName("Точка нулевого тока " + pgOrigName);
			sheetParamsData.Columns(1).SetUnits("В");
			sheetParamsData.Columns(2).SetLongName("Ионный ток насыщения " + pgOrigName);
			sheetParamsData.Columns(2).SetUnits("А");
			sheetParamsData.Columns(3).SetLongName("Температура " + pgOrigName);
			sheetParamsData.Columns(3).SetUnits("эВ");
			sheetParamsData.Columns(4).SetLongName("Электронная плотность " + pgOrigName);
			sheetParamsData.Columns(4).SetUnits("см^-3");
			break;
	}
	
	Tree trFormat;  
	trFormat.Root.CommonStyle.Fill.FillColor.nVal = 18; 
	DataRange dr;
	for (int i = 1; i < DIM1 + 1; ++i)
	{
		string time;
		
		time.Format("%.4f - %.4f", vTime[i-1], vTime[i-1] + 1.0 / freqP - 0.0001);
		strUnits.Format("%i - %i", vStartSegIndxs[i-1],
			(i < DIM1) ? vStartSegIndxs[i] - 1 : vStartSegIndxs[i-1] + vStartSegIndxs[i-1] - vStartSegIndxs[i-2] - 1);
		
		sheetOrigData.Columns(i).SetUnits(strUnits);
		sheetOrigData.Columns(i).SetLongName(time);
		sheetOrigData.Columns(i).SetComments((string)i);
		
		sheetFitData.Columns(i).SetUnits(strCutOffs);
		sheetFitData.Columns(i).SetLongName(time);
		sheetFitData.Columns(i).SetComments((string)i);
		
		if (diagnostics == 1)
		{
			sheetDiffData.Columns(i).SetUnits(strUnits);
			sheetDiffData.Columns(i).SetLongName(time);
			sheetDiffData.Columns(i).SetComments((string)i);
		}
		
		if (filtS != 0)
		{
			sheetFiltData.Columns(i).SetUnits(strFilt);
			sheetFiltData.Columns(i).SetLongName(time);
			sheetFiltData.Columns(i).SetComments((string)i);
		}
		
		if (i % 2 != 0)
		{
			dr.Add("X", sheetOrigData, 0, i, -1, i);
			if (0==dr.UpdateThemeIDs(trFormat.Root)) bool bRet = dr.ApplyFormat(trFormat, true, true);
			dr.Reset();
			
			dr.Add("X", sheetFitData, 0, i, -1, i);
			if (0==dr.UpdateThemeIDs(trFormat.Root)) bool bRet = dr.ApplyFormat(trFormat, true, true);
			dr.Reset();
			
			if (diagnostics == 1)
			{	
				dr.Add("X", sheetDiffData, 0, i, -1, i);
				if (0==dr.UpdateThemeIDs(trFormat.Root)) bool bRet = dr.ApplyFormat(trFormat, true, true);
				dr.Reset();
			}
			
			dr.Add("X", sheetParamsData, i, 0, i, -1);
			if (0==dr.UpdateThemeIDs(trFormat.Root)) bool bRet = dr.ApplyFormat(trFormat, true, true);
			dr.Reset();
			
			if(filtS != 0)
			{
				dr.Add("X", sheetFiltData, 0, i, -1, i);
				if (0==dr.UpdateThemeIDs(trFormat.Root)) bool bRet = dr.ApplyFormat(trFormat, true, true);
				dr.Reset();
			}
		}
	}

	s_currp = pg.GetName();
	
	return pg;
}

void compare_books(uint column,
				   vector <string> & vResultBooksNames)
{
	int i, j, is_ICRH = 0;
	
	/* create book */
	Worksheet sheetCompare, sheetOrigData;
	sheetCompare.Create("Origin");
	sheetCompare.SetSize(-1, 3);
	sheetCompare.Columns(1).SetLongName("c ИЦР");
	sheetCompare.Columns(2).SetLongName("без ИЦР");
	
	string scol, s_beg, s_end;
	
	switch (column)
	{
	case 0:
		scol = "A";
		break;
	case 1:
		scol = "B";
		break;
	case 2:
		scol = "C";
		break;
	case 3:
		scol = "D";
		break;
	case 4:
		scol = "E";
		break;
	default:
		return;
	}
	
	Worksheet sheet = Project.FindPage(vResultBooksNames[0]).Layers(Project.FindPage(vResultBooksNames[0]).Layers.Count() - 1);
	sheetCompare.GetPage().SetLongName(FindPageNameNumber("Сравнение " +
			get_str_from_str_list(Project.FindPage(vResultBooksNames[0]).GetLongName(), " ", 0) + 
			" " + sheet.Columns(column).GetLongName()));
	
	for (i = 0; i < vResultBooksNames.GetSize(); ++i)
	{
		Page p = Project.FindPage(vResultBooksNames[i]);
		Worksheet sheetData = p.Layers(p.Layers.Count() - 1);
		
		/* Берем книгу с данными эксперимента, которая прикреплена к книге результата обработки */
		vector <uint> vUIDs;
		if (p.Layers(0).GetConnectedObjects(vUIDs) <= 0)
		{
			MessageBox(GetWindow(), "Не найдена книга с оригинальными данными");
			return;
		}
		
		for (j = 0; j < vUIDs.GetSize(); ++j)
		{
			sheetOrigData = Project.GetObject(vUIDs[j]);
			if (sheetOrigData)
				break;
		}
		
		if (!sheetOrigData)
		{
			MessageBox(GetWindow(), "Недействительна книга с оригинальными данными");
			return;
		}
		
		vector vX, vY;
		vY = ColIdxByName(sheetOrigData, "ИЦР пр").GetDataObject();
		vX.SetSize(vY.GetSize());
		
		if (max(vY) > 0.5) // есть ИЦР
		{
			is_ICRH = 1;
			
			vector <uint> vecIndex, vBuff(4);
			
			{ // блок обработки сигнала ИЦР генератора
				for (j = 0; j < vX.GetSize(); ++j)
					vX[j] = j;
				
				/* Берем производную от столбика с ИЦР, чтобы найти начало и конец включения ИЦР генератора */
				ocmath_derivative(vX, vY, vY.GetSize());
				
				vY.Find(vecIndex, max(vY)); // находит индекс максимального значения и кладет его в vecIndex
				vBuff[0] = vecIndex[0];
				vY.Find(vecIndex, min(vY)); // находит индекс минимального значения и кладет его в vecIndex
				vBuff[1] = vecIndex[0];
				
				if (vBuff[0] > vBuff[1])
					swap_i((int)vBuff[0], (int)vBuff[1]);
			}
			
			{ // блок обработки сигнала ВЧ генератора
				vY = ColIdxByName(sheetOrigData, "ВЧ Пр").GetDataObject();
				vX.SetSize(vY.GetSize());
				
				for (j = 0; j < vX.GetSize(); ++j)
					vX[j] = j;
				
				/* Берем производную от столбика с ВЧ, чтобы найти начало и конец включения ВЧ генератора */
				ocmath_derivative(vX, vY, vY.GetSize());
				
				vY.Find(vecIndex, max(vY)); // находит индекс максимального значения и кладет его в vecIndex
				vBuff[2] = vecIndex[0];
				vY.Find(vecIndex, min(vY)); // находит индекс минимального значения и кладет его в vecIndex
				vBuff[3] = vecIndex[0];
				
				if (vBuff[2] > vBuff[3])
					swap_i((int)vBuff[2], (int)vBuff[3]);
			}
			
			s_beg = (int)((double)(vBuff[0] - vBuff[2]) / (vBuff[3] - vBuff[2]) * sheetData.Columns(column).GetUpperBound());
			s_end = (int)((double)(vBuff[1] - vBuff[2]) / (vBuff[3] - vBuff[2]) * sheetData.Columns(column).GetUpperBound());
		}
		
		/* Записываем название книги обработки в первую колонку */
		sheetCompare.SetCell(i, 0, get_str_from_str_list(p.GetLongName(), " ", 3));
		
		if (is_ICRH == 1) // есть ИЦР
		{
			/* Записываем усредненные значения нужного праметра во время ИЦР во вторую колонку */
			//"=mean([Book25]Sheet1!A)"
			string str = "=mean([" + p.GetName() + "]" + sheetData.GetName() + 
				"!" + scol + "[" + s_beg + ":" + s_end + "])";
			sheetCompare.SetCell(i, 1, str);
			
			/* Записываем усредненные значения нужного праметра без ИЦР в третью колонку */
			//"=(mean([Book25]Sheet1!A[1:2]) + mean([Book25]Sheet1!A[4:5])) / 2"
			str = "=(mean([" + p.GetName() + "]" + sheetData.GetName() + 
				"!" + scol + "[1:" + s_beg + "])" +
				" + mean([" + p.GetName() + "]" + sheetData.GetName() + 
				"!" + scol + "[" + s_end + ":0])) / 2";
			sheetCompare.SetCell(i, 2, str);
		}
		else 
		{
			/* Записываем усредненные значения нужного праметра без ИЦР во третью колонку */
			//"=mean([Book25]Sheet1!A)"
			string str = "=mean([" + p.GetName() + "]" + sheetData.GetName() + 
				"!" + scol + ")";
			sheetCompare.SetCell(i, 2, str);
		}
	}
	
	
}