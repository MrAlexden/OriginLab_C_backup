#include "AllinOneScript_rework.h"

int swap_graph(int diagnostics, int swap_koeff)
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
		end; // Индекс конца отрезка

	/* Читаем необходимые параметры */
	is_str_numeric_integer(get_str_from_str_list(sheetLegend.Columns(1).GetLongName(), " ", 1), &begin);
	is_str_numeric_integer(get_str_from_str_list(sheetLegend.Columns(1).GetLongName(), " ", 3), &end);
	
	vI[0] = ColIdxByUnit(sheetOrigData, (string)begin).GetIndex() + swap_koeff;
	
	if (vI[0] == 0 && swap_koeff == -1 
		|| vI[0] == sheetOrigData.GetNumCols() && swap_koeff == 1)
	{
		MessageBox(GetWindow(), "Это крайний график!");
		return -1;
	}
	
	switch (diagnostics)
	{
		case 0:
			
			/* Блок заполнения графика */
			if (fill_legend(pgLegend,
							pgData,
							vI) < 0)
				return -1;
			
			if (make_add_data_for_graph_zond(pgLegend,
											 pgData,
											 vI) < 0)
				return -1;
			
			GraphPage gp = Project.ActiveLayer().GetPage();
			
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
			
		case 1:
			
			/* Блок заполнения графика */
			if (fill_legend(pgLegend,
							pgData,
							vI) < 0)
				return -1;
			
			if (make_add_data_for_graph_setka(pgLegend,
											  pgData,
											  vI) < 0)
				return -1;
			
			GraphPage gp = Project.ActiveLayer().GetPage();
			
			if (fill_single_segment_graph(gp,
										  pgData,
										  vI) < 0)
				return -1;
			
			if (add_sublines_to_graph_setka(gp,
										   pgLegend) < 0)
				return -1;
			
			create_labels(1);
			/* Конец блока заполнения графика */
			
			return 0;
			
		case 2:
			
			/* Блок заполнения графика */
			if (fill_legend(pgLegend,
							pgData,
							vI) < 0)
				return -1;
			
			if (make_add_data_for_graph_cilin(pgLegend,
											  pgData,
											  vI) < 0)
				return -1;
			
			GraphPage gp = Project.ActiveLayer().GetPage();
			
			if (fill_single_segment_graph(gp,
										  pgData,
										  vI) < 0)
				return -1;
			
			if (add_sublines_to_graph_cilin(gp,
										   pgLegend) < 0)
				return -1;
			
			create_labels(2);
			/* Конец блока заполнения графика */
			
			return 0;
			
		case 3:
			
			/* Блок заполнения графика */
			if (fill_legend(pgLegend,
							pgData,
							vI) < 0)
				return -1;
			
			if (make_add_data_for_graph_doubleprobe(pgLegend,
													pgData,
													vI) < 0)
				return -1;
			
			GraphPage gp = Project.ActiveLayer().GetPage();
			
			if (fill_single_segment_graph(gp,
										  pgData,
										  vI) < 0)
				return -1;
				
			if (add_sublines_to_graph_doubleprobe(gp,
												  pgLegend) < 0)
			return -1;
			
			create_labels(3);
			/* Конец блока заполнения графика */
			
			return 0;
	}
	
	return 0;
}