#include "AllinOneScript_rework.h"

GraphPage create_empty_graph(Page & pgData,
							 vector <int> & vI)
{
	GraphPage gp;
	gp.Create("Origin");
	GraphLayer gl1 = gp.Layers(0);
	pgData.Layers(0).ConnectTo(gl1);
	Tree tr;
	tr = gl1.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
	tr.Root.Legend.Font.Size.nVal = 12;
	gl1.UpdateThemeIDs(tr.Root);
	gl1.ApplyFormat(tr, true, true);
	
	gl1.XAxis.Additional.ZeroLine.nVal = 1;
	gl1.YAxis.Additional.ZeroLine.nVal = 1;
	
	tr.Reset();
	tr.Root.Grids.HorizontalMajorGrids.AddNode("Show").nVal = 1, 
	tr.Root.Grids.VerticalMajorGrids.AddNode("Show").nVal = 1;
	tr.Root.Grids.HorizontalMajorGrids.Color.nVal = 18;
	tr.Root.Grids.HorizontalMajorGrids.Style.nVal = 5; 
	tr.Root.Grids.HorizontalMajorGrids.Width.dVal = 1;
	tr.Root.Grids.VerticalMajorGrids.Color.nVal = 18;
	tr.Root.Grids.VerticalMajorGrids.Style.nVal = 5; 
	tr.Root.Grids.VerticalMajorGrids.Width.dVal = 1;
	
    gl1.YAxis.UpdateThemeIDs(tr.Root);
    
    gl1.XAxis.ApplyFormat(tr, true, true);
    gl1.YAxis.ApplyFormat(tr, true, true);
    
	if (pgData.GetLongName().Match("Сетка*"))
		page_add_layer(gp, false, false, false, true, ADD_LAYER_INIT_SIZE_POS_MOVE_OFFSET, false, 0, LINK_STRAIGHT);
	
	if (vI.GetSize() > 1)
		gp.SetLongName(FindPageNameNumber(get_str_from_str_list(pgData.GetLongName(), " ", 0)
				  + " " + get_str_from_str_list(pgData.GetLongName(), " ", 3)
				  + '(' + get_str_from_str_list(pgData.GetLongName(), " ", 4)
				  + ')' + " Отрезки " + (vI[0]) + " - " + (vI[vI.GetSize() - 1])));
	
	return gp;
}

int take_selection(vector <int> & vI, int & comp_evo)
{
	Worksheet sheetCurrent = Project.ActiveLayer(),
			  sheetOrigData = WksIdxByName(sheetCurrent.GetPage(), "OrigData"),
			  sheetFitData = WksIdxByName(sheetCurrent.GetPage(), "Fit"),
			  sheetParamsData = WksIdxByName(sheetCurrent.GetPage(), "Parameters"),
			  sheetFiltData = WksIdxByName(sheetCurrent.GetPage(), "Filtration"),
			  sheetDiffData = WksIdxByName(sheetCurrent.GetPage(), "Diff");

	if (!sheetOrigData 
		|| !sheetFitData 
		|| !sheetParamsData)
    {
    	MessageBox(GetWindow(), "Ошибка! Не найден лист с данными");
		return -1;
    }		  
    
    int beginRow,
        endRow,
        beginCol,
        endCol,
        i = 0,
        j = 0;
    
    /* Берем выбранный диапазон */
    if(!sheetCurrent.GetSelectedRange(beginRow, beginCol, endRow, endCol))
	{
		if(sheetCurrent == sheetParamsData) 
			MessageBox(GetWindow(), "!Selection (выберите строку)");
		else 
			MessageBox(GetWindow(), "!Selection (выберите колонку)");
		return -1;
	}
	
	/* Заполняем вектор с индексами выбранных сегментов */
	if(sheetCurrent == sheetParamsData)
	{
		vI.SetSize(abs(endRow - beginRow) + 1);
		for (i = 0, j = beginRow; j < endRow + 1; ++i, ++j)
			vI[i] = j + 1;
	}
	else /* vI будет дейстителен на всех листах кроме Params, там vI[]-1 */
	{
		if(beginCol == 0) 
			beginCol = 1;
		if(endCol == 0) 
			endCol = 1;
		
		vI.SetSize(abs(endCol - beginCol) + 1);
		for (i = 0, j = beginCol; j < endCol + 1; ++i, ++j)
			vI[i] = j;
	}
	
	/* Проверка ошибки */
	vector v = sheetParamsData.Columns(0).GetDataObject();
    if(v[vI[vI.GetSize() - 1] - 1] == 0)
	{
		MessageBox(GetWindow(), "Ошибка во времени! Последний отрезок обработан с ошибкой (время начала равно 0). "
			+ "Для решения проблемы удалите отрезки со временем начала 0, или пересчитайте импульс");
		return -1;
	}
    
	/* Диалог если выбрано несколько отрезков */
    if(vI.GetSize() > 1)
	{
		GETN_BOX( treeTest ); 
		GETN_RADIO_INDEX_EX(decayT1, "", 0, "Сравнить выбранные отрезки|Построить эволюцию выбранных отрезков") 
		
		if(0==GetNBox( treeTest, "График нескольких отрезков" )) 
			return-1;
		
		treeTest.GetNode("decayT1").GetValue(comp_evo);
	}
	else
		comp_evo = 1; /* Мини костыль для того чтобы не создавать второй лист в книге легенды */
	
	return 0;
}

Page & create_legend(Page & pgData)
{
	Worksheet wks;
	wks.Create("Table", CREATE_HIDDEN);

	GraphLayer gl = Project.ActiveLayer();	
	GraphObject grTable = gl.CreateLinkTable(pgData.GetLongName(), wks);
	wks.ConnectTo(gl);
	
	return wks.GetPage();
}

int fill_legend(Page & pgLegend,
				Page & pgData,
				vector <int> & vI)
{
	Worksheet sheetLegend = pgLegend.Layers(0),
			  sheetOrigData = WksIdxByName(pgData, "OrigData"),
			  sheetParamsData = WksIdxByName(pgData, "Parameters");
	
	string NameOfGraphPage,
		   strImpulseType,
		   strFuelType,
		   strSegTime,
		   strSegPoints;
	
	strImpulseType = get_str_from_str_list(sheetOrigData.Columns(0).GetUnits(), " ", 0);
	strFuelType = get_str_from_str_list(sheetOrigData.Columns(0).GetUnits(), " ", 1);
 
	if(vI.GetSize() > 1)
	{
		NameOfGraphPage = get_str_from_str_list(pgData.GetLongName(), " ", 0)
						  + " " + get_str_from_str_list(pgData.GetLongName(), " ", 3)
						  + '(' + get_str_from_str_list(pgData.GetLongName(), " ", 4)
						  + ')' + " Отрезки " + (vI[0]) + " - " + (vI[vI.GetSize() - 1]);
		
		strSegTime = get_str_from_str_list(sheetOrigData.Columns(min(vI)).GetLongName(), " ", 0) + " - " + 
				get_str_from_str_list(sheetOrigData.Columns(max(vI)).GetLongName(), " ", 2);
		strSegPoints = get_str_from_str_list(sheetOrigData.Columns(min(vI)).GetUnits(), " ", 0) + " - " + 
				get_str_from_str_list(sheetOrigData.Columns(max(vI)).GetUnits(), " ", 2);
		
		sheetLegend.Columns(0).SetLongName("Отрезок времени " + strSegTime);
		sheetLegend.Columns(1).SetComments("Значения");
		sheetLegend.Columns(0).SetComments(NameOfGraphPage); 
		sheetLegend.Columns(1).SetLongName("Строки "+ strSegPoints);
		
		sheetLegend.SetCell(0, 0, "Тип импульса");
		sheetLegend.SetCell(0, 1, strImpulseType);
		
		sheetLegend.SetCell(1, 0, "Тип топлива");
		sheetLegend.SetCell(1, 1, strFuelType);
	}
	else
	{
		NameOfGraphPage = get_str_from_str_list(pgData.GetLongName(), " ", 0)
						  + " " + get_str_from_str_list(pgData.GetLongName(), " ", 3)
						  + '(' + get_str_from_str_list(pgData.GetLongName(), " ", 4)
						  + ')' + " Отрезок " + (vI[0]);
		
		strSegTime = sheetOrigData.Columns(vI[0]).GetLongName();
		strSegPoints = sheetOrigData.Columns(vI[0]).GetUnits();
		
		sheetLegend.Columns(0).SetLongName("Отрезок времени " + strSegTime);
		sheetLegend.Columns(1).SetComments("Значения");
		sheetLegend.Columns(0).SetComments(NameOfGraphPage); 
		sheetLegend.Columns(1).SetLongName("Строки "+ strSegPoints);
		
		vector v;
		
		v = sheetParamsData.Columns(1).GetDataObject();
		sheetLegend.SetCell(0, 0, sheetParamsData.Columns(1).GetLongName() + ", " +
								  sheetParamsData.Columns(1).GetUnits());
		sheetLegend.SetCell(0, 1, v[vI[0] - 1]);
		
		v = sheetParamsData.Columns(2).GetDataObject();
		sheetLegend.SetCell(1, 0, sheetParamsData.Columns(2).GetLongName() + ", " +
								  sheetParamsData.Columns(2).GetUnits());
		sheetLegend.SetCell(1, 1, v[vI[0] - 1]);
		
		v = sheetParamsData.Columns(3).GetDataObject();
		sheetLegend.SetCell(2, 0, sheetParamsData.Columns(3).GetLongName() + ", " +
								  sheetParamsData.Columns(3).GetUnits());
		sheetLegend.SetCell(2, 1, v[vI[0] - 1]);
		
		v = sheetParamsData.Columns(4).GetDataObject();
		sheetLegend.SetCell(3, 0, sheetParamsData.Columns(4).GetLongName() + ", " +
								  sheetParamsData.Columns(4).GetUnits());
		sheetLegend.SetCell(3, 1, v[vI[0] - 1]);
		
		sheetLegend.SetCell(4, 0, "Тип импульса");
		sheetLegend.SetCell(4, 1, strImpulseType);
		
		sheetLegend.SetCell(5, 0, "Тип топлива");
		sheetLegend.SetCell(5, 1, strFuelType);
	}
	
	sheetLegend.AutoSize();
	GraphPage gp = Project.ActiveLayer().GetPage();
	gp.Layers(0).GraphObjects(pgData.GetLongName()).AutoSize();
	gp.Layers(0).GraphObjects(pgData.GetLongName()).Width = 2500;
	
	return 0;
}

int fill_single_segment_graph(GraphPage & gp,
							  Page & pgData,
							  vector <int> & vI)
{
	/* Нужно для того, чтобы работало и на сетке и на цилиндре */
	int LayerNum = gp.Layers(1) ? 1 : 0;
	
	Worksheet sheetOrigData = WksIdxByName(pgData, "OrigData"),
			  sheetFitData = WksIdxByName(pgData, "Fit"),
			  sheetParamsData = WksIdxByName(pgData, "Parameters"),
			  sheetFiltData = WksIdxByName(pgData, "Filtration"),
			  sheetDiffData = WksIdxByName(pgData, "Diff");

	while (gp.Layers(0).RemovePlot(0));
	if (gp.Layers(1))
		while (gp.Layers(1).RemovePlot(0));
 
	/* Original Data sheet */
	Curve cOrigData(sheetOrigData, 0 , vI[0]);
	gp.Layers(0).AddPlot(cOrigData, IDM_PLOT_LINE);
	
	/* Differentiated Data sheet */
	if (sheetDiffData) /* Если есть производная */
	{
		Curve cDiffData(sheetDiffData, 0 , vI[0]);
		gp.Layers(LayerNum).AddPlot(cDiffData, IDM_PLOT_LINE);
	}
	
#ifdef GAUSSFIT
	/* Fit Data sheet */
	Curve cFitData(sheetFitData, 0 , vI[0]);
	gp.Layers(LayerNum).AddPlot(cFitData, IDM_PLOT_LINE);
#else
	/* Fit Data sheet */
	Curve cFitData(sheetFitData, 0 , vI[0]);
	gp.Layers(0).AddPlot(cFitData, IDM_PLOT_LINE);
#endif
	
	/* Filtrated Data sheet */
	Curve cFiltData(sheetFiltData, 0 , vI[0]);
	if (sheetFiltData && sheetFiltData.Columns(vI[0]).GetUpperBound() > 0) /* Если есть фильтрация */
		gp.Layers(0).AddPlot(cFiltData, IDM_PLOT_LINE);
	
	DataPlot dpOrigData = gp.Layers(0).DataPlots(0),
			 dpFitData,
			 dpFiltData,
			 dpDiffData;

#ifdef GAUSSFIT
	if (!gp.Layers(1))
		dpFitData = gp.Layers(0).DataPlots(1);
	else
	{
		dpDiffData = gp.Layers(1).DataPlots(0);
		dpFitData = gp.Layers(1).DataPlots(1);
	}
#else
	dpFitData = gp.Layers(0).DataPlots(1);
	if (gp.Layers(1))
		dpDiffData = gp.Layers(1).DataPlots(0);
#endif
	
	if (sheetFiltData && sheetFiltData.Columns(vI[0]).GetUpperBound() > 0)
		dpFiltData = gp.Layers(0).DataPlots(gp.Layers(0).DataPlots.Count() - 1);
	
	Tree trOrigData, trFitData;
	trOrigData = dpOrigData.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
	trFitData = dpFitData.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
	
	trOrigData.Root.Line.Width.dVal = 2;
	trOrigData.Root.Line.Color.nVal = set_color_rnd();
	
	trFitData.Root.Line.Width.dVal = 2.5;
	trFitData.Root.Line.Color.nVal = 1;
	
	dpOrigData.ApplyFormat(trOrigData, true, true);
	dpFitData.ApplyFormat(trFitData, true, true);
	
	if (sheetFiltData && sheetFiltData.Columns(vI[0]).GetUpperBound() > 0)
	{
		Tree trFiltData;
		trFiltData = dpFiltData.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
		
		trFiltData.Root.Line.Width.dVal = 2.5;
		trFiltData.Root.Line.Color.nVal = 2;
		
		dpFiltData.ApplyFormat(trFiltData, true, true);
		
		vector <uint> vecUI = {0, 2, 1};
#ifdef GAUSSFIT
		if (!gp.Layers(1))
			gp.Layers(0).ReorderPlots(vecUI);
#else
		gp.Layers(0).ReorderPlots(vecUI);
#endif
	}
	
	if (sheetDiffData)
	{
		Tree trDiffData;
		trDiffData = dpDiffData.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
		
		trDiffData.Root.Line.Width.dVal = 2.5;
		trDiffData.Root.Line.Color.nVal = 3;
		
		dpDiffData.ApplyFormat(trDiffData, true, true);
	}
	
	gp.Layers(0).Rescale();
	if (gp.Layers(1)) /* Нужно для того, чтобы работало и на сетке и на цилиндре */
		gp.Layers(1).Rescale();
	
	vector vPila;
	vPila = sheetOrigData.Columns(0).GetDataObject();
	vPila.SetSize(sheetOrigData.Columns(vI[0]).GetUpperBound());
	
	if (vPila[0] >= 0) 
		gp.Layers(0).XAxis.Scale.From.dVal = 0;
	else 
		gp.Layers(0).XAxis.Scale.From.dVal = vPila[0];
	
	gp.Layers(0).XAxis.Scale.To.dVal = vPila[vPila.GetSize()-1];
	gp.Layers(0).XAxis.Scale.IncrementBy.nVal = 0; // 0=increment by value; 1=number of major ticks
	gp.Layers(0).XAxis.Scale.Value.dVal = 20; // Increment value
	
	gp.Layers(0).YAxis.Scale.IncrementBy.nVal = 1; // 0=increment by value; 1=number of major ticks
	gp.Layers(0).YAxis.Scale.MajorTicksCount.dVal = 10;
	
	if (gp.Layers(1)) /* Нужно для того, чтобы работало и на сетке и на цилиндре */
	{
		gp.Layers(1).YAxis.Scale.IncrementBy.nVal = 1; // 0=increment by value; 1=number of major ticks
		gp.Layers(1).YAxis.Scale.MajorTicksCount.dVal = 10;
	}
	
	gp.SetLongName(FindPageNameNumber(get_str_from_str_list(pgData.GetLongName(), " ", 0)
					  + " " + get_str_from_str_list(pgData.GetLongName(), " ", 3)
					  + '(' + get_str_from_str_list(pgData.GetLongName(), " ", 4)
					  + ')' + " Отрезок " + (vI[0])));
	
	return 0;
}

int fill_compare_graph(GraphPage & gp,
					   vector <int> & vI,
					   Page & pgData)
{
	/* Нужно для того, чтобы работало и на сетке и на цилиндре */
	int LayerNum = gp.Layers(1) ? 1 : 0;
	
	Worksheet sheetOrigData = WksIdxByName(pgData, "OrigData"),
			  sheetFitData = WksIdxByName(pgData, "Fit");
	
	Tree tr;
	Curve cc;
	int i = 0, k = 0, color = 0;
	
	/* ORIGDATA */
	for(i = 0; i < vI.GetSize(); ++i, ++k)
	{
		cc.Attach(sheetOrigData, 0, vI[i]);
		gp.Layers(0).AddPlot(cc, IDM_PLOT_SCATTER);
		
		tr = gp.Layers(0).DataPlots(k).GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
		tr.Root.Symbol.Size.nVal = 2;
		gp.Layers(0).DataPlots(k).ApplyFormat(tr, true, true);
		color = set_color_except_bad(color);
		gp.Layers(0).DataPlots(k).SetColor(color);
	}
	
	if (gp.Layers(1)) /* Если сетка, где 2 масштаба */
		k = 0;
	
	/* FITTING */
	for(i = 0; i < vI.GetSize(); ++i, ++k)
	{
		cc.Attach(sheetFitData, 0, vI[i]);
		
		gp.Layers(LayerNum).AddPlot(cc, IDM_PLOT_LINE);
		
		tr = gp.Layers(LayerNum).DataPlots(k).GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
		
		tr.Root.Line.Width.dVal = 2;
		color = set_color_except_bad(color);
		
		gp.Layers(LayerNum).DataPlots(k).ApplyFormat(tr, true, true);
		gp.Layers(LayerNum).DataPlots(k).SetColor(color);
	}
	
	gp.Layers(0).Rescale();
	if (gp.Layers(1)) /* Нужно для того, чтобы работало и на сетке и на цилиндре */
		gp.Layers(1).Rescale();
	
	vector vPila;
	vPila = sheetOrigData.Columns(0).GetDataObject();
	
	if (vPila[0] >= 0) 
		gp.Layers(0).XAxis.Scale.From.dVal = 0;
	else 
		gp.Layers(0).XAxis.Scale.From.dVal = vPila[0];
	
	gp.Layers(0).XAxis.Scale.To.dVal = vPila[vPila.GetSize()-1];
	gp.Layers(0).XAxis.Scale.IncrementBy.nVal = 0; // 0=increment by value; 1=number of major ticks
	gp.Layers(0).XAxis.Scale.Value.dVal = 20; // Increment value
	
	gp.Layers(0).YAxis.Scale.IncrementBy.nVal = 1; // 0=increment by value; 1=number of major ticks
	gp.Layers(0).YAxis.Scale.MajorTicksCount.dVal = 10;
	
	if (gp.Layers(1)) /* Нужно для того, чтобы работало и на сетке и на цилиндре */
	{
		gp.Layers(1).YAxis.Scale.IncrementBy.nVal = 1; // 0=increment by value; 1=number of major ticks
		gp.Layers(1).YAxis.Scale.MajorTicksCount.dVal = 10;
	}
	
	gp.SetLongName(FindPageNameNumber(get_str_from_str_list(pgData.GetLongName(), " ", 0)
					  + " " + get_str_from_str_list(pgData.GetLongName(), " ", 3)
					  + '(' + get_str_from_str_list(pgData.GetLongName(), " ", 4)
					  + ')' + " Отрезки " + (vI[0]) + " - " + (vI[vI.GetSize() - 1])));
	
	return 0;
}

int fill_evolution_graph(GraphPage & gp,
					     Page & pgLegend)
{
	/* Нужно для того, чтобы работало и на сетке и на цилиндре */
	int LayerNum = gp.Layers(1) ? 1 : 0;
	
	Worksheet wks = pgLegend.Layers(1);
   	
	/* Original Data */
	Curve cOrigData(wks, 0, 1);
	gp.Layers(0).AddPlot(cOrigData, IDM_PLOT_LINE);
	
	/* Fit Data */
	Curve cFitData(wks, 0, 2);
	gp.Layers(LayerNum).AddPlot(cFitData, IDM_PLOT_LINE);
	
	/* Filtrated Data */
	if(wks.Columns(3).GetUpperBound() > 0)
	{
		Curve cFiltData(wks, 0, 3);
		gp.Layers(0).AddPlot(cFiltData, IDM_PLOT_LINE);
	}
	
	DataPlot dpOrigData = gp.Layers(0).DataPlots(0),
			 dpFitData, 
			 dpFiltData;
	
	if (!gp.Layers(1)) /* Если зонд или цилиндр, где все в одном масштабе */
		dpFitData = gp.Layers(0).DataPlots(1);
	else /* Если сетка, где 2 масштаба */
		dpFitData = gp.Layers(1).DataPlots(0);
	
	if (wks.Columns(3).GetUpperBound() > 0)
		dpFiltData = gp.Layers(0).DataPlots(gp.Layers(0).DataPlots.Count() - 1);

	Tree trOrigData, trFitData, trFiltData;
	
	trOrigData = dpOrigData.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
	trFitData = dpFitData.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
	if(wks.Columns(3).GetUpperBound() > 0) 
		trFiltData = dpFiltData.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
	
	trOrigData.Root.Line.Width.dVal = 1;
	trOrigData.Root.Line.Color.nVal = 0;
	
	trFitData.Root.Line.Width.dVal = 2;
	trFitData.Root.Line.Color.nVal = 1;
	
	trFiltData.Root.Line.Width.dVal = 2;
	trFiltData.Root.Line.Color.nVal = 2;
	
	dpOrigData.ApplyFormat(trOrigData, true, true);
	dpFitData.ApplyFormat(trFitData, true, true);
	if (wks.Columns(3).GetUpperBound() > 0)
	{
		dpFiltData.ApplyFormat(trFiltData, true, true);
		
		vector <uint> vecUI = {0, 2, 1};
		if (!gp.Layers(1))
			gp.Layers(0).ReorderPlots(vecUI);
	}
	
	gp.Layers(0).Rescale();
	if (gp.Layers(1)) /* Нужно для того, чтобы работало и на сетке и на цилиндре */
		gp.Layers(1).Rescale();
	
	gp.Layers(0).XAxis.Scale.IncrementBy.nVal = 1; // 0=increment by value; 1=number of major ticks
	gp.Layers(0).XAxis.Scale.MajorTicksCount.dVal = 10;
	gp.Layers(0).YAxis.Scale.IncrementBy.nVal = 1; // 0=increment by value; 1=number of major ticks
	gp.Layers(0).YAxis.Scale.MajorTicksCount.dVal = 10;
	
	return 0;
}

void create_labels(int z_s_c)
{
	GraphPage gp = Project.ActiveLayer().GetPage();
	GraphLayer gl1 = gp.Layers(0), gl2= gp.Layers(1), gl;
	if(gl2) gl = gl2;
	else gl = gl1;
	
	Page pgData, /* Книга с результатами обработки */
		 pgOrigData, /* Книга с оригинальными данными (TDMS) */
		 pgLegend; /* Книга с данными для легенды (table) */
	
	Get_OriginData_wks(pgData, pgOrigData, pgLegend);
	
	Worksheet wks_sheetLegend = pgLegend.Layers(0), 
			  wks_reflines = pgLegend.Layers(1);
	
	Dataset slave, slave1;
	Curve crv;
	Tree tr;
	tr = gl.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
	
	switch (z_s_c)
	{
	case 0:
		if(wks_reflines.GetNumCols() <= 4)
			wks_reflines.SetSize(-1, 7);
		
		slave.Attach(wks_reflines.Columns(4));
		slave.SetSize(2);
		slave1.Attach(wks_reflines.Columns(0));
		slave[0] = slave1[1];
		slave1.Attach(wks_reflines.Columns(2));
		slave[1] = slave1[0];
		
		slave.Attach(wks_reflines.Columns(5));
		slave.SetSize(2);
		slave1.Attach(wks_reflines.Columns(1));
		slave[0] = slave1[1];
		slave1.Attach(wks_reflines.Columns(3));
		slave[1] = slave1[0];
		
		slave.Attach(wks_reflines.Columns(6));
		slave.SetSize(2);
		slave1.Attach(wks_reflines.Columns(2));
		slave.SetText(0,"A+B*x");
		slave[1] = slave1[0];
		
		wks_reflines.Columns(6).SetType(OKDATAOBJ_DESIGNATION_L);
		crv.Attach(wks_reflines,4,5);
		gl.AddLabelPlot(crv, wks_reflines.Columns(6));
		
		break;
		
	case 3:
		if(wks_reflines.GetNumCols() <= 4)
			wks_reflines.SetSize(-1, 7);
		
		slave.Attach(wks_reflines.Columns(4));
		slave.SetSize(2);
		slave1.Attach(wks_reflines.Columns(0));
		slave[0] = slave1[1];
		slave1.Attach(wks_reflines.Columns(2));
		slave[1] = slave1[0];
		
		slave.Attach(wks_reflines.Columns(5));
		slave.SetSize(2);
		slave1.Attach(wks_reflines.Columns(1));
		slave[0] = slave1[1];
		slave1.Attach(wks_reflines.Columns(3));
		slave[1] = slave1[0];
		
		slave.Attach(wks_reflines.Columns(6));
		slave.SetSize(2);
		slave1.Attach(wks_reflines.Columns(1));
		slave[0] = slave1[1];
		slave1.Attach(wks_reflines.Columns(3));
		slave[1] = slave1[0];
		
		wks_reflines.Columns(6).SetType(OKDATAOBJ_DESIGNATION_L);
		crv.Attach(wks_reflines,4,5);
		gl.AddLabelPlot(crv, wks_reflines.Columns(6));
		
		break;
		
	default:
		if(wks_reflines.GetNumCols()<=7) wks_reflines.SetSize(-1,10);
		
		slave.Attach(wks_reflines.Columns(7));
		slave.SetSize(4);
		slave1.Attach(wks_reflines.Columns(0));
		slave[0] = slave[1] = slave1[1];
		slave1.Attach(wks_reflines.Columns(3));
		slave[2] = slave1[0];
		slave1.Attach(wks_reflines.Columns(5));
		slave[3] = slave1[0];
		
		slave.Attach(wks_reflines.Columns(8));
		slave.SetSize(4);
		slave1.Attach(wks_reflines.Columns(1));
		slave[0] = slave1[0];
		slave1.Attach(wks_reflines.Columns(2));
		slave[1] = slave1[0];
		slave1.Attach(wks_reflines.Columns(4));
		slave[2] = slave[3] = slave1[0];
		
		slave.Attach(wks_reflines.Columns(9));
		slave.SetSize(4);
		slave1.Attach(wks_reflines.Columns(8));
		slave[0] = slave1[0];
		slave[1] = slave1[1];
		slave1.Attach(wks_reflines.Columns(7));
		slave[2] = slave1[2];
		slave[3] = slave1[3];
		
		wks_reflines.Columns(9).SetType(OKDATAOBJ_DESIGNATION_L);
		crv.Attach(wks_reflines,7,8);
		gl.AddLabelPlot(crv, wks_reflines.Columns(9));
		
		break;
	}
    
	tr.Root.DataLabels.DataLabel1.Size.nVal = 12;
	tr.Root.DataLabels.DataLabel1.Color.nVal = 3;
	tr.Root.DataLabels.DataLabel1.YOffset.nVal = 40;
	tr.Root.DataLabels.DataLabel1.XOffset.nVal = 40;
	if (0==gl.UpdateThemeIDs(tr.Root)) gl.ApplyFormat(tr, true, true);
}

void rezult_graphs(Page & pgData, int z_s_c)
{
	string pgOrigName = pgData.GetLongName();
	Worksheet wks_sheet = WksIdxByName(pgData, "Parameters");
	
	if(z_s_c == 0 || z_s_c == 3)
	{
		GraphPage gpTemp, gpDens;
		gpTemp.Create("Origin");
		gpDens.Create("Origin");
		GraphLayer glTemp = gpTemp.Layers(0), glDens = gpDens.Layers(0);
		gpTemp.SetLongName ( FindPageNameNumber ( get_str_from_str_list(pgOrigName, " ", 0) + " " + get_str_from_str_list(pgOrigName, " ", 3) + '(' + get_str_from_str_list(pgOrigName, " ", 4) + ')' + " Температура" ) );
		gpDens.SetLongName ( FindPageNameNumber ( get_str_from_str_list(pgOrigName, " ", 0) + " " + get_str_from_str_list(pgOrigName, " ", 3) + '(' + get_str_from_str_list(pgOrigName, " ", 4) + ')' + " Плотность" ) );
		Tree tr;
		tr = glTemp.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
		tr.Root.Legend.Font.Size.nVal = 12;
		if(0 == glTemp.UpdateThemeIDs(tr.Root) ) 
		{
			glTemp.ApplyFormat(tr, true, true);
			glDens.ApplyFormat(tr, true, true);
		}
		
		DataRange drTemp, drDens;
		DataPlot dpTemp, dpDens;
		tr.Reset();
		
		drTemp.Add(wks_sheet, 0, "X");
		drTemp.Add(wks_sheet, 3, "Y");
		
		drDens.Add(wks_sheet, 0, "X");
		drDens.Add(wks_sheet, 4, "Y");
		
		glTemp.AddPlot(drTemp, IDM_PLOT_LINESYMB);
		glDens.AddPlot(drDens, IDM_PLOT_LINESYMB);
		
		dpTemp = glTemp.DataPlots(0);
		dpDens = glDens.DataPlots(0);
		
		tr = dpTemp.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
		tr.Root.Line.Width.dVal=1.1;
		tr.Root.Line.Color.nVal=0;
		tr.Root.Symbol.Size.nVal=5;;
		tr.Root.Symbol.EdgeColor.nVal=0;
		
		dpTemp.ApplyFormat(tr, true, true);
		dpDens.ApplyFormat(tr, true, true);
		
		glTemp.Rescale();
		glDens.Rescale();
		
		Axis axisX = glTemp.XAxis, axisY = glTemp.YAxis;	
		axisX.Additional.ZeroLine.nVal = 1;
		axisY.Additional.ZeroLine.nVal = 1;

		//tr.Reset();
		
		// Show major grid
		TreeNode trProperty = tr.Root.Grids.HorizontalMajorGrids.AddNode("Show"), trProperty1 = tr.Root.Grids.VerticalMajorGrids.AddNode("Show");
		trProperty.nVal = 1;
		trProperty1.nVal = 1;
		tr.Root.Grids.HorizontalMajorGrids.Color.nVal = 18;
		tr.Root.Grids.HorizontalMajorGrids.Style.nVal = 5; 
		tr.Root.Grids.HorizontalMajorGrids.Width.dVal = 1;
		tr.Root.Grids.VerticalMajorGrids.Color.nVal = 18;
		tr.Root.Grids.VerticalMajorGrids.Style.nVal = 5; 
		tr.Root.Grids.VerticalMajorGrids.Width.dVal = 1;
		if(0 == axisY.UpdateThemeIDs(tr.Root)) 
		{
			axisY.ApplyFormat(tr, true, true);
			axisX.ApplyFormat(tr, true, true);
		}	
		axisX.Scale.From.dVal = 0;
		axisX.Scale.IncrementBy.nVal = 1; // 0=increment by value; 1=number of major ticks
		axisX.Scale.MajorTicksCount.dVal = 10;
		axisY.Scale.From.dVal = 0;
		axisY.Scale.IncrementBy.nVal = 1; // 0=increment by value; 1=number of major ticks
		axisY.Scale.MajorTicksCount.dVal = 10;
		
		axisX = glDens.XAxis, axisY = glDens.YAxis;
		axisX.Additional.ZeroLine.nVal = 1;
		axisY.Additional.ZeroLine.nVal = 1;
		if(0 == axisY.UpdateThemeIDs(tr.Root)) 
		{
			axisY.ApplyFormat(tr, true, true);
			axisX.ApplyFormat(tr, true, true);
		}
		axisX.Scale.From.dVal = 0;
		axisX.Scale.IncrementBy.nVal = 1; // 0=increment by value; 1=number of major ticks
		axisX.Scale.MajorTicksCount.dVal = 10;
		axisY.Scale.From.dVal = 0;
		axisY.Scale.IncrementBy.nVal = 1; // 0=increment by value; 1=number of major ticks
		axisY.Scale.MajorTicksCount.dVal = 10;
		
		glTemp.GraphObjects("XB").Text = "Время, с";
		glTemp.GraphObjects("YL").Text = "Температура электронов, эВ";
		glDens.GraphObjects("XB").Text = "Время, с";
		glDens.GraphObjects("YL").Text = "Плотность плазмы, см^-3";
		
		set_active_layer(glDens);
		
		legend_update(glTemp, ALM_LONGNAME);
		legend_update(glDens, ALM_LONGNAME);
	}
	else
	{
		GraphPage gpTemp, gpPV, gpEnergy;
		gpTemp.Create("Origin");
		gpPV.Create("Origin");
		//if(z_s_c == 2) gpEnergy.Create("Origin");
		GraphLayer glTemp = gpTemp.Layers(0), glPV = gpPV.Layers(0), glEnergy = gpEnergy.Layers(0);
		
		gpTemp.SetLongName ( FindPageNameNumber ( get_str_from_str_list(pgOrigName, " ", 0) + " " + get_str_from_str_list(pgOrigName, " ", 3) + '(' + get_str_from_str_list(pgOrigName, " ", 4) + ')' + " Температура" ) );
		gpPV.SetLongName ( FindPageNameNumber ( get_str_from_str_list(pgOrigName, " ", 0) + " " + get_str_from_str_list(pgOrigName, " ", 3) + '(' + get_str_from_str_list(pgOrigName, " ", 4) + ')' + " Потенциал плазмы" ) );
		//gpEnergy.SetLongName ( FindPageNameNumber ( get_str_from_str_list(pgOrigName, " ", 0) + " " + get_str_from_str_list(pgOrigName, " ", 3) + '(' + get_str_from_str_list(pgOrigName, " ", 4) + ')' + " Энергия" ) );
		
		Tree tr;
		tr = glTemp.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
		tr.Root.Legend.Font.Size.nVal = 12;
		if(0 == glTemp.UpdateThemeIDs(tr.Root) ) 
		{
			glTemp.ApplyFormat(tr, true, true);
			glPV.ApplyFormat(tr, true, true);
			//if(z_s_c == 2) glEnergy.ApplyFormat(tr, true, true);
		}
	
		DataRange drTemp, drPV, drEnergy;
		DataPlot dpTemp, dpPV, dpEnergy;
		tr.Reset();
		
		drTemp.Add(wks_sheet, 0, "X");
		drTemp.Add(wks_sheet, 2, "Y");
		
		drPV.Add(wks_sheet, 0, "X");
		drPV.Add(wks_sheet, 3, "Y");
		
		//if(z_s_c == 2) 
		//{
			//drEnergy.Add(wks_sheet, 0, "X");
			//drEnergy.Add(wks_sheet, 4, "Y");
		//}
		
		glTemp.AddPlot(drTemp, IDM_PLOT_LINESYMB);
		glPV.AddPlot(drPV, IDM_PLOT_LINESYMB);
		//if(z_s_c == 2) glEnergy.AddPlot(drEnergy, IDM_PLOT_LINESYMB);
		
		dpTemp = glTemp.DataPlots(0);
		dpPV = glPV.DataPlots(0);
		//if(z_s_c == 2) dpEnergy = glEnergy.DataPlots(0);
		
		tr = dpTemp.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
		tr.Root.Line.Width.dVal=1.1;
		tr.Root.Line.Color.nVal=0;
		tr.Root.Symbol.Size.nVal=5;;
		tr.Root.Symbol.EdgeColor.nVal=0;
		
		dpTemp.ApplyFormat(tr, true, true);
		dpPV.ApplyFormat(tr, true, true);
		//if(z_s_c == 2) dpEnergy.ApplyFormat(tr, true, true);
		
		glTemp.Rescale();
		glPV.Rescale();
		//if(z_s_c == 2) glEnergy.Rescale();
		
		Axis axisX = glTemp.XAxis, axisY = glTemp.YAxis;	
		axisX.Additional.ZeroLine.nVal = 1;
		axisY.Additional.ZeroLine.nVal = 1;

		//tr.Reset();
		
		// Show major grid
		TreeNode trProperty = tr.Root.Grids.HorizontalMajorGrids.AddNode("Show"), trProperty1 = tr.Root.Grids.VerticalMajorGrids.AddNode("Show");
		trProperty.nVal = 1;
		trProperty1.nVal = 1;
		tr.Root.Grids.HorizontalMajorGrids.Color.nVal = 18;
		tr.Root.Grids.HorizontalMajorGrids.Style.nVal = 5; 
		tr.Root.Grids.HorizontalMajorGrids.Width.dVal = 1;
		tr.Root.Grids.VerticalMajorGrids.Color.nVal = 18;
		tr.Root.Grids.VerticalMajorGrids.Style.nVal = 5; 
		tr.Root.Grids.VerticalMajorGrids.Width.dVal = 1;
		if(0 == axisY.UpdateThemeIDs(tr.Root)) 
		{
			axisY.ApplyFormat(tr, true, true);
			axisX.ApplyFormat(tr, true, true);
		}	
		axisX.Scale.From.dVal = 0;
		axisX.Scale.IncrementBy.nVal = 1; // 0=increment by value; 1=number of major ticks
		axisX.Scale.MajorTicksCount.dVal = 10;
		axisY.Scale.From.dVal = 0;
		axisY.Scale.IncrementBy.nVal = 1; // 0=increment by value; 1=number of major ticks
		axisY.Scale.MajorTicksCount.dVal = 10;
		
		axisX = glPV.XAxis, axisY = glPV.YAxis;
		axisX.Additional.ZeroLine.nVal = 1;
		axisY.Additional.ZeroLine.nVal = 1;
		if(0 == axisY.UpdateThemeIDs(tr.Root)) 
		{
			axisY.ApplyFormat(tr, true, true);
			axisX.ApplyFormat(tr, true, true);
		}
		axisX.Scale.From.dVal = 0;
		axisX.Scale.IncrementBy.nVal = 1; // 0=increment by value; 1=number of major ticks
		axisX.Scale.MajorTicksCount.dVal = 10;
		axisY.Scale.From.dVal = 0;
		axisY.Scale.IncrementBy.nVal = 1; // 0=increment by value; 1=number of major ticks
		axisY.Scale.MajorTicksCount.dVal = 10;
		
		//if(z_s_c == 2) 
		//{
			//axisX = glEnergy.XAxis, axisY = glEnergy.YAxis;
			//axisX.Additional.ZeroLine.nVal = 1;
			//axisY.Additional.ZeroLine.nVal = 1;
			//if(0 == axisY.UpdateThemeIDs(tr.Root)) 
			//{
				//axisY.ApplyFormat(tr, true, true);
				//axisX.ApplyFormat(tr, true, true);
			//}
			//axisX.Scale.From.dVal = 0;
			//axisX.Scale.IncrementBy.nVal = 1; // 0=increment by value; 1=number of major ticks
			//axisX.Scale.MajorTicksCount.dVal = 10;
			////axisY.Scale.From.dVal = 0;
			//axisY.Scale.IncrementBy.nVal = 1; // 0=increment by value; 1=number of major ticks
			//axisY.Scale.MajorTicksCount.dVal = 10;
		//}
		
		glTemp.GraphObjects("XB").Text = "Время, с";
		glTemp.GraphObjects("YL").Text = "Температура ионов, эВ";
		glPV.GraphObjects("XB").Text = "Время, с";
		glPV.GraphObjects("YL").Text = "Потенциал плазмы, В";
		//if(z_s_c == 2) 
		//{
			//glEnergy.GraphObjects("XB").Text = "Время, с";
			//glEnergy.GraphObjects("YL").Text = "Энергия, эВ";
		//}
		
		set_active_layer(glTemp);
		
		legend_update(glTemp, ALM_LONGNAME);
		legend_update(glPV, ALM_LONGNAME);
	}	
}

int set_color_rnd()
{
	time_t ltime;
	time( &ltime );
	int color = ltime % 23;
	while (color == 1
		|| color == 2
		|| color == 3
		|| color == 4
		|| color == 6
		|| color == 17
		|| color == 20
		|| color == 21)
	{
		color = rand() % 23;
	}
	return color;
}

int set_color_except_bad(int incol)
{
	switch (incol)
	{
		case 0:
			return ++incol;
			break;
		case 1:
			return ++incol;
			break;
		case 2:
			return ++incol;
			break;
		case 3:
			return incol+=2;
			break;
		case 4:
			return ++incol;
			break;
		case 5:
			return incol+=2;
			break;
		case 6:
			return ++incol;
			break;
		case 7:
			return ++incol;
			break;
		case 8:
			return ++incol;
			break;
		case 9:
			return ++incol;
			break;
		case 10:
			return ++incol;
			break;
		case 11:
			return ++incol;
			break;
		case 12:
			return ++incol;
			break;
		case 13:
			return ++incol;
			break;
		case 14:
			return ++incol;
			break;
		case 15:
			return ++incol;
			break;
		case 16:
			return incol+=2;
			break;
		case 17:
			return ++incol;
			break;
		case 18:
			return ++incol;
			break;
		case 19:
			return ++incol;
			break;
		case 20:
			return incol+=3;
			break;
		case 21:
			return incol+=2;
			break;
		case 22:
			return ++incol;
			break;
		case 23:
			return 0;
			break;
		default:
			return 0;
	}
	return 0;
}