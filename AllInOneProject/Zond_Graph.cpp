#include "AllinOneScript_rework.h"

int make_add_data_for_graph_zond(Page & pgLegend,
								 Page & pgData,
								 vector <int> & vI)
{
	if (!pgLegend)
		return -1;
	
	if (!pgLegend.Layers(1))
		if (vI.GetSize() > 1) 
			pgLegend.AddLayer("SeveralColumns");
		else
			pgLegend.AddLayer("A+B*x");
	
	Worksheet wksAdd = pgLegend.Layers(1),
			  sheetOrigData = WksIdxByName(pgData, "OrigData"),
			  sheetFitData = WksIdxByName(pgData, "Fit"),
			  sheetParamsData = WksIdxByName(pgData, "Parameters"),
			  sheetFiltData = WksIdxByName(pgData, "Filtration");

	while(wksAdd.DeleteCol(0));
	wksAdd.SetSize(-1,4);
	
	vector v;
	
	if (vI.GetSize() == 1)
	{
		wksAdd.Columns(1).SetLongName("A+B*x");
		wksAdd.Columns(3).SetLongName("Пересечение с OX");
		
		Dataset dsX, dsY;
		
		v = sheetOrigData.Columns(0).GetDataObject(); /* Pila */
		
		dsX.Attach(wksAdd.Columns(0));
		dsX.SetSize(2);
		dsX[0] = v[0];
		dsX[1] = v[v.GetSize()-1];
		
		linfit(sheetFitData.Columns(0).GetDataObject(), sheetFitData.Columns(vI[0]).GetDataObject(), 0.3, 1, v);
		
		/* v[0] -> A, v[1] -> B */
		dsY.Attach(wksAdd.Columns(1));
		dsY.SetSize(2);
		dsY[0] = v[0] + dsX[0] * v[1];
		dsY[1] = v[0] + dsX[1] * v[1];
		
		v = sheetParamsData.Columns(1).GetDataObject(); /* Parameters([1] -> flo pot) */
		
		dsX.Attach(wksAdd.Columns(2));
		dsX.SetSize(2);
		dsX[0] = v[vI[0] - 1];
		dsX[1] = v[vI[0] - 1];
		
		v = sheetFitData.Columns(vI[0]).GetDataObject(); /* OrigData */
		
		dsY.Attach(wksAdd.Columns(3));
		dsY.SetSize(2);
		dsY[0] = min(v);
		dsY[1] = max(v);
	}
	else
	{
		wksAdd.Columns(1).SetLongName("Сигнал");
		wksAdd.Columns(2).SetLongName("Аппроксимация");
		wksAdd.Columns(3).SetLongName("Фильтрация");
		
		int i = 0;
		vector vTime, 
			   vOrigData,
			   vFiltData,
			   vFitData;
		
		/* Заполняем вектора с данными */
		for(i = 0; i < vI.GetSize(); ++i)
		{
			vOrigData.Append(sheetOrigData.Columns(vI[i]).GetDataObject());
			
			vFitData.Append(sheetFitData.Columns(vI[i]).GetDataObject());
			
			if (sheetFiltData)
			{
				if(sheetFiltData.Columns(vI[i]).GetUpperBound() <= 0)
					vFiltData.SetSize(vFiltData.GetSize() + sheetFiltData.Columns(0).GetUpperBound());
				else
					vFiltData.Append(sheetFiltData.Columns(vI[i]).GetDataObject());
			}
		}
		
		/* Заполняем вектор со временем */ 
		vTime.SetSize(vOrigData.GetSize());
		v = sheetParamsData.Columns(0).GetDataObject();
		double dt = (v[vI[vI.GetSize()-1]] - v[vI[0]]) / vTime.GetSize();
		for (i = 1, vTime[0] = v[vI[0] - 1]; i < vTime.GetSize(); ++i)
			vTime[i] = vTime[i - 1] + dt;
		
		Dataset ds;
		
		ds.Attach(wksAdd.Columns(0));
		ds = vTime;
		
		ds.Attach(wksAdd.Columns(1));
		ds = vOrigData;
		
		ds.Attach(wksAdd.Columns(2));
		ds = vFitData;
		
		ds.Attach(wksAdd.Columns(3));
		ds = vFiltData;
	}
	
	return 0;
}

int add_sublines_to_graph_zond(GraphPage & gp,
							   Page & pgLegend)
{
	Worksheet wks = pgLegend.Layers(1);
	
	/* A + B*x line */
	Curve cABX(wks, 0, 1);
	gp.Layers(0).AddPlot(cABX, IDM_PLOT_LINE);
	
	/* Floating potential line */
	Curve cFloP(wks, 2, 3);
	gp.Layers(0).AddPlot(cFloP, IDM_PLOT_LINE);
	
	DataPlot dpABX = gp.Layers(0).DataPlots(gp.Layers(0).DataPlots.Count() - 2),
			 dpFloP = gp.Layers(0).DataPlots(gp.Layers(0).DataPlots.Count() - 1);
	
	Tree trAdd;
	trAdd = dpABX.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
	
	trAdd.Root.Line.Width.dVal = 1.1;
	trAdd.Root.Line.Color.nVal = 3;
	
	dpABX.ApplyFormat(trAdd, true, true);
	dpFloP.ApplyFormat(trAdd, true, true);
	
	legend_update(gp.Layers(0), ALM_LONGNAME);
	
	return 0;
}

int graphZond()
{	
	Page pgData = Project.ActiveLayer().GetPage();
    if (!pgData)
    {
    	MessageBox(GetWindow(), "book!");
    	return -1;
    }

	vector <int> vI;
	int numSegs, 
		comp_evo, 
		i = 0, 
		j = 0;

	if (take_selection(vI, comp_evo) < 0)
		return -1;
	
	numSegs = vI.GetSize();

	Worksheet sheetOrigData = WksIdxByName(pgData, "OrigData"),
			  sheetFitData = WksIdxByName(pgData, "Fit"),
			  sheetParamsData = WksIdxByName(pgData, "Parameters"),
			  sheetFiltData = WksIdxByName(pgData, "Filtration");
	
	GraphPage gp = create_empty_graph(pgData, vI);
	
	Page pgLegend = create_legend(pgData);
	
	if (fill_legend(pgLegend,
				    pgData,
				    vI) < 0)
        return -1;

    /* if single segment and evolution(begin of next seg connected to the end og previous) */
	if (comp_evo == 1)
		if (make_add_data_for_graph_zond(pgLegend,
										 pgData,
										 vI) < 0)
		return -1;
    
	/* if single segment */
    if (numSegs == 1)
    {
		if (fill_single_segment_graph(gp,
									  pgData,
									  vI) < 0)
			return -1;
		
		if (add_sublines_to_graph_zond(gp,
							           pgLegend) < 0)
			return -1;
    }
    
	/* if comparison (all segments each one under another) */
    if (comp_evo == 0 && numSegs > 1)
		if (fill_compare_graph(gp,
							   vI,
							   pgData) < 0)
			return -1;
    
	/* if evolution(begin of next seg connected to the end og previous) */
    if(comp_evo == 1 && numSegs > 1)
		if (fill_evolution_graph(gp,
							     pgLegend) < 0)
		return -1;

	if(comp_evo == 0) 
		gp.Layers(0).GraphObjects("XB").Text = "Время, с";
	else 
		gp.Layers(0).GraphObjects("XB").Text = "Напряжение, В";
	gp.Layers(0).GraphObjects("YL").Text = "Ток, А";
	
	legend_update(gp.Layers(0), ALM_LONGNAME);
    
    if(numSegs == 1)
		create_labels(0);
    
    return 0;
}

int reprocess_single_segment_zond(int begin,
								  int end,
								  int colnum,
								  Worksheet & sheetOrigData,
			                      Worksheet & sheetFitData,
			                      Worksheet & sheetParamsData,
			                      Worksheet & sheetFiltData,
			                      Worksheet & sheetLegend, 
			                      Page & pgData,
			                      Page & pgOrigData)
{
	Dataset ds;
	int i = 0, j = 0;
	
	int origSize = end - begin + 1;
    int curSize = sheetFitData.Columns(colnum).GetUpperBound() + 1;
    int newSize = origSize * (1 - leftP - rightP);
	
	vector vPila(sheetOrigData.Columns(0)), 
		   vSignal(sheetOrigData.Columns(colnum)), 
		   vAdd(6),
		   vResultData(newSize), 
		   vFiltedData(newSize),
		   vDiffedData(newSize),
		   vParameters(4);

 	string strFuelType;
    sheetLegend.GetCell(5, 1, strFuelType);
    
    vAdd[0] = S;
	vAdd[1] = linfitP;
	vAdd[2] = filtS;
	vAdd[3] = (strFuelType.Match("He")) ? 0 : 1;
	vAdd[4] = Num_iter;
	vAdd[5] = newSize;
	
	/* Заполняем вектора нового размера, которые будем обрабатывать */
    if (newSize <= curSize)
    {
		vPila.SetSize(newSize);
		vSignal.SetSize(newSize);
		
		sheetOrigData.Columns(colnum).SetNumRows(newSize);
		sheetFitData.Columns(colnum).SetNumRows(newSize);
    }
    else
    {
    	double step = vPila[vPila.GetSize() - 1] - vPila[vPila.GetSize() - 2];
    	
    	vPila.SetSize(newSize);
		vSignal.SetSize(newSize);
		
    	for (i = curSize; i < newSize; ++i)
    		vPila[i] = vPila[i-1] + step;
    	
    	Worksheet wksOrigData = pgOrigData.Layers(0);
    	Column col = ColIdxByName(wksOrigData, get_str_from_str_list(sheetParamsData.Columns(0).GetComments(), "|", 1));
    	if (!col) 
    	{
    		MessageBox(GetWindow(), ERR_GetErrorDescription(-222));
    		return -1;
    	}
    	
    	vector vOrigSignal(col);
    	
    	for (i = 0, j = begin + origSize * leftP; i < newSize; ++i, ++j)
    		vSignal[i] = vOrigSignal[j] / -resistance;
    	
    	/* Вставляю новые пилы в книгу только если они больше старых */
    	if (newSize > sheetOrigData.Columns(0).GetUpperBound() + 1)
    	{
			ds.Attach(sheetOrigData.Columns(0));
			ds = vPila;
			ds.Attach(sheetFitData.Columns(0));
			ds = vPila;
			if (sheetFiltData && filtS != 0)
			{
				ds.Attach(sheetFiltData.Columns(0));
				ds = vPila;
			}
    	}
    }
    
	int err = OriginMakeOne(0, vPila, vSignal, vAdd, vResultData, vFiltedData, vDiffedData, vParameters);
	if (err < 0)
		return err;
	
	/* Original Data sheet */
	ds.Attach(sheetOrigData.Columns(colnum));
	ds.SetSize(newSize);
	ds = vSignal;
	
	/* Fit Data sheet */
	ds.Attach(sheetFitData.Columns(colnum));
	ds.SetSize(newSize);
	ds = vResultData;
	
	/* Create Filtrated Data sheet if it doesn't exist and filtrating was used */
	if (!sheetFiltData && filtS != 0)
	{
		pgData.AddLayer("Filtration");
		sheetFiltData = pgData.Layers(pgData.Layers.Count() - 1);
		
		while (sheetFiltData.DeleteCol(0));
		sheetFiltData.SetSize(sheetOrigData.Columns(0).GetUpperBound(), sheetOrigData.Columns.Count());
		
		sheetFiltData.Columns(0).SetComments("0");
		sheetFiltData.Columns(0).SetType(OKDATAOBJ_DESIGNATION_X);
		sheetFiltData.Columns(0).SetUnits(sheetOrigData.Columns(0).GetUnits());
		sheetFiltData.Columns(0).SetLongName("pila");
		
		ds.Attach(sheetFiltData.Columns(0));
		ds.SetSize(newSize);
		ds = sheetOrigData.Columns(0).GetDataObject();
		
		Tree trFormat;  
		trFormat.Root.CommonStyle.Fill.FillColor.nVal = 18; 
		DataRange dr;
		for (i = 1; i < sheetOrigData.Columns.Count(); ++i)
		{
			if (i % 2 != 0)
			{
				if(filtS != 0)
				{
					dr.Add("X", sheetFiltData, 0, i, -1, i);
					if (0==dr.UpdateThemeIDs(trFormat.Root)) bool bRet = dr.ApplyFormat(trFormat, true, true);
					dr.Reset();
				}
			}
		}
		
		sheetFiltData.SetIndex(1);
	}
	
	/* if there was filtrating and Filtrated Data sheet alredy exists */
	if (filtS != 0)
	{
		ds.Attach(sheetFiltData.Columns(colnum));
		ds.SetSize(newSize);
		ds = vFiltedData;
		
		sheetFiltData.Columns(colnum).SetNumRows(newSize);
		
		string strFilt;
		strFilt.Format("%.2f", filtS);
		sheetFiltData.Columns(colnum).SetUnits(strFilt);
	}
	
	/* if there was no filtrating */
	if (sheetFiltData && filtS == 0)
	{
		ds.Attach(sheetFiltData.Columns(colnum));
		ds.SetSize(0);
		
		sheetFiltData.Columns(colnum).SetUnits((string)0);
	}

	/* Parameters */
	ds.Attach(sheetParamsData.Columns(1));
	ds[colnum - 1] = vParameters[0];
	
	ds.Attach(sheetParamsData.Columns(2));
	ds[colnum - 1] = vParameters[1];
	
	ds.Attach(sheetParamsData.Columns(3));
	ds[colnum - 1] = vParameters[2];
	
	ds.Attach(sheetParamsData.Columns(4));
	ds[colnum - 1] = vParameters[3];
	
	string strCutOffs;
    strCutOffs.Format("B %.2f E %.2f LF %.2f", leftP, rightP, linfitP);
	sheetFitData.Columns(colnum).SetUnits(strCutOffs);
	
	return 0;
}

int reprocess_zond()
{
	Page pgData, /* Книга с результатами обработки */
		 pgOrigData, /* Книга с оригинальными данными (TDMS) */
		 pgLegend; /* Книга с данными для легенды (table) */
	
	if (Get_OriginData_wks(pgData, pgOrigData, pgLegend) < 0)
		return -1;
	
	Worksheet sheetOrigData = WksIdxByName(pgData, "OrigData"),
			  sheetFitData = WksIdxByName(pgData, "Fit"),
			  sheetParamsData = WksIdxByName(pgData, "Parameters"),
			  sheetFiltData = WksIdxByName(pgData, "Filtration"),
			  sheetLegend = pgLegend.Layers(0);
	
	if (sheetLegend.Columns(0).GetUpperBound() <= 2)
	{
		MessageBox(GetWindow(), "Невозможно пересчитать, на графике больше одного отрезка");
		return -1;
	}
	
	vector <int> vI(1);	
	int begin, // Индекс начала отрезка
		end, // Индекс конца отрезка
		isFilted,
		i = 0,
		j = 0;
	
	/* Читаем необходимые параметры */
	is_str_numeric_integer(get_str_from_str_list(sheetLegend.Columns(1).GetLongName(), " ", 1), &begin);
	is_str_numeric_integer(get_str_from_str_list(sheetLegend.Columns(1).GetLongName(), " ", 3), &end);
	is_str_numeric_integer(get_str_from_str_list(sheetParamsData.Columns(0).GetUnits(), " ", 2), &resistance);
	is_str_numeric(get_str_from_str_list(sheetParamsData.Columns(0).GetUnits(), " ", 3), &S);
	
	vI[0] = ColIdxByUnit(sheetOrigData, (string)begin).GetIndex();
	
	is_str_numeric(get_str_from_str_list(sheetFitData.Columns(vI[0]).GetUnits(), " ", 1), &leftP);
    is_str_numeric(get_str_from_str_list(sheetFitData.Columns(vI[0]).GetUnits(), " ", 3), &rightP);
    is_str_numeric(get_str_from_str_list(sheetFitData.Columns(vI[0]).GetUnits(), " ", 5), &linfitP);
    
    /* СОЗДАНИЕ ДИАЛОГА */
	GETN_BOX( treeTest ); 

    GETN_SLIDEREDIT(right, "Часть точек, которые будут отсечены от конца", rightP, "0.01|0.89|88")
    
	GETN_SLIDEREDIT(linfit, "Часть данных от начала которые будут линейно аппроксимированны", linfitP, "0.00|0.9|90")
	
	GETN_NUM(iter, "Колическтво итераций метода Левенберга-Марквардта", Num_iter)

	if (sheetFiltData) 
		if (sheetFiltData.Columns(vI[0]).GetUpperBound() > 0)
		{
			is_str_numeric(sheetFiltData.Columns(vI[0]).GetUnits(), &filtS);
			isFilted = 1;
		}
		else
			isFilted = 0;
    else
    	isFilted = 0;

    GETN_BEGIN_BRANCH(FiltOpt, "Фильтрация данных") GETN_CHECKBOX_BRANCH(isFilted)
    if (isFilted == 1) 
    	GETN_OPTION_BRANCH(GETNBRANCH_OPEN)
		GETN_SLIDEREDIT(filt, "Часть точек фильтрации", filtS, "0.00|0.99|99") 
    GETN_END_BRANCH(FiltOpt)
    
    GETN_BEGIN_BRANCH(UseSame, "Применить к нескольким отрезкам") GETN_CHECKBOX_BRANCH(0)
		string str;
		str.Format("1|%i|%i", sheetOrigData.Columns.Count() - 1, sheetOrigData.Columns.Count() - 2);
		GETN_SLIDEREDIT(B, "От:", vI[0], str)
		GETN_SLIDEREDIT(E, "До:", sheetOrigData.Columns.Count() - 1, str)
	GETN_END_BRANCH(UseSame)
    
	if (0 == GetNBox( treeTest, "Пересчет этого отрезка" ))
		return -1;
	/* КОНЕЦ СОЗДАНИЯ ДИАЛОГА */
	
	/* СЧИТЫВАНИЕ ДАННЫХ ИЗ ДИАЛОГА */
	treeTest.GetNode("right").GetValue(rightP);
	
	treeTest.GetNode("linfit").GetValue(linfitP);
	
	treeTest.GetNode("iter").GetValue(Num_iter);
	
	if (treeTest.FiltOpt.Use)
		treeTest.GetNode("FiltOpt").GetNode("filt").GetValue(filtS);
	else 
		filtS = 0;
	
	if (treeTest.UseSame.Use)
	{
		treeTest.GetNode("UseSame").GetNode("B").GetValue(i);
		treeTest.GetNode("UseSame").GetNode("E").GetValue(j);
		
		if (i > j)
		{
			MessageBox(GetWindow(), "Индекс первого отрезка больше индекса последнего");
			return -1;
		}
		
		vI.SetSize(j - i + 1);
		for (j = 0; j < vI.GetSize(); ++j, ++i)
			vI[j] = i;
	}
	/* КОНЕЦ СЧИТЫВАНИЯ ДАННЫХ ИЗ ДИАЛОГА */
	
	progressBox prgbBox("Применение изменений");
	prgbBox.SetRange(0, vI.GetSize());
	
	for (i = 0; i < vI.GetSize(); ++i)
	{
		if (!prgbBox.Set(i)) return -404;
		
		is_str_numeric_integer(get_str_from_str_list(sheetOrigData.Columns(vI[i]).GetUnits(), " ", 0), &begin);
		is_str_numeric_integer(get_str_from_str_list(sheetOrigData.Columns(vI[i]).GetUnits(), " ", 2), &end);
		
		int err = reprocess_single_segment_zond(begin,
										        end,
										        vI[i],
										        sheetOrigData,
										        sheetFitData,
										        sheetParamsData,
										        sheetFiltData,
										        sheetLegend, 
										        pgData,
										        pgOrigData);
		if (err < 0)
			return err;
	}
	
	/* Подрезаем размеры листов по уровню колонок */
    sheetOrigData.GetBounds(begin, 1, end, -1, false);
    sheetOrigData.SetSize(end + 1, -1);
    sheetFitData.SetSize(end + 1, -1);
    if (sheetFiltData) 
    	sheetFiltData.SetSize(end + 1, -1);
    
    /* Блок заполнения графика */
    is_str_numeric_integer(get_str_from_str_list(sheetLegend.Columns(1).GetLongName(), " ", 1), &begin);
	is_str_numeric_integer(get_str_from_str_list(sheetLegend.Columns(1).GetLongName(), " ", 3), &end);
	vI.SetSize(1);
	vI[0] = ColIdxByUnit(sheetOrigData, (string)begin).GetIndex();
    
    if (fill_legend(pgLegend,
				    pgData,
				    vI) < 0)
        return -1;
    
    if (make_add_data_for_graph_zond(pgLegend,
									 pgData,
									 vI) < 0)
		return -1;

    GraphPage gp = Project.ActiveLayer().GetPage();
    while (gp.Layers(0).RemovePlot(0));
    if (fill_single_segment_graph(gp,
								  pgData,
								  vI) < 0)
		return -1;
	
	if (add_sublines_to_graph_zond(gp,
								   pgLegend) < 0)
		return -1;
	
	create_labels(0);
	/* Конец блока заполнения графика */
	
	return 0;
}