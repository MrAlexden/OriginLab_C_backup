#include "AllinOneScript_rework.h"

int linfit(vector & vX, 
		   vector & vY, 
		   double linfitP, 
		   int direction, // 1 -- слева, -1 -- справа
		   vector & vParams)
{		
	int i, j;
	vector vXhandler, 
		   vYhandler;
	vXhandler = vX; 
	vYhandler = vY;	 
	if (direction == 1)
	{
		vXhandler.SetSize(ceil(vXhandler.GetSize() * linfitP));
		vYhandler.SetSize(ceil(vYhandler.GetSize() * linfitP));
	}
	if (direction == -1)
	{
		vXhandler.SetSize(ceil(vXhandler.GetSize() * linfitP));
		for (i = 0, j = vX.GetSize() * (1 - linfitP); i < vXhandler.GetSize();)
			vXhandler[i++] = vX[j++];
		vYhandler.SetSize(ceil(vYhandler.GetSize() * linfitP));
		for (i = 0, j = vY.GetSize() * (1 - linfitP); i < vYhandler.GetSize();)
			vYhandler[i++] = vY[j++];
	}
	vParams.SetSize(2);
	
	LROptions stLROptions;
	RegStats stRegStats;
	RegANOVA stRegANOVA;
	FitParameter psFitParameter[2];	 
	
	ocmath_linear_fit(vXhandler, vYhandler, vYhandler.GetSize(), psFitParameter, NULL, 0, &stLROptions, &stRegStats, &stRegANOVA);
	
	vParams[0] = psFitParameter[0].Value;
	vParams[1] = psFitParameter[1].Value;
	
	return 0;
}

bool is_signalpeakslookingdown(vector & v)
{
    double min = min(v),
		   max = max(v),
		   vavg,
		   s1,
		   s2;
	
	vector vh;
	
	v.GetSubVector(vh, 0, v.GetSize() * 0.1);
	
	vh.Sum(s1);
	
	v.GetSubVector(vh, v.GetSize() * 0.9, v.GetSize() - 1);
	
	vh.Sum(s2);
	
    vavg = (s1 + s2) / (v.GetSize() * 0.2);

    if (abs(vavg - max) >= abs(vavg - min))
        return false;
    else
        return true;
}

int data_process(int diagnostics,		 // diagnostics type (zond::0|setka::1|cilind::2)
				 string DiagnosticName,
				 Worksheet & wks, 
			     Column & colPila, 
			     Column & colSignal)
{
	if(!colSignal) return -222;
	if(!colPila) return -333;
	
	int DIM1,	// DIM1 - numsegs
		DIM2,	// DIM2 - sizeseg
		i = 0,
		j = 0,
		k = 0,
		err = 0,
		buffer = 15000; 
	vector vPila(colPila), 
		   vSignal(colSignal), 
		   vAdd(14), 
		   vResP(buffer),
		   vOriginalData, 
		   vResultData, 
		   vFiltedData,
		   vDiffedData,
		   vParameters;
	vector <int> vStartSegIndxs(buffer);

	
	vAdd[0] = S;
	vAdd[1] = st_time_end_time[0];
	vAdd[2] = st_time_end_time[1];
	vAdd[3] = leftP;
	vAdd[4] = rightP;
	vAdd[5] = linfitP;
	vAdd[6] = filtS;
	vAdd[7] = freqP;
	vAdd[8] = resistance;
	vAdd[9] = coefPila;
	vAdd[10] = FuelType;
	vAdd[11] = Num_iter;
	vAdd[12] = vPila.GetSize();
	vAdd[13] = vSignal.GetSize();
	
	err = OriginFindSignal(diagnostics, vPila, vSignal, vAdd, buffer, &DIM1, &DIM2, vResP, vStartSegIndxs);
	if (err < 0)
		return err;

	vResP.SetSize(DIM2);
	vStartSegIndxs.SetSize(DIM1);
	
	vOriginalData.SetSize(DIM1 * DIM2);
	vResultData.SetSize(DIM1 * DIM2);
	vFiltedData.SetSize(DIM1 * DIM2);
	vDiffedData.SetSize(DIM1 * DIM2);
	vParameters.SetSize(DIM1 * Number_of_rezult_plasma_parametrs);
	
	err = OriginAll(diagnostics, vPila, vSignal, vAdd, vOriginalData, vResultData, vFiltedData, vDiffedData, vParameters);
	if (err < 0)
		return err;
	
	vector v(DIM1);
	for (i = 0, j = 0; j < DIM1; i += Number_of_rezult_plasma_parametrs, ++j)
		v[j] = vParameters[i];
	
	Page pg = create_book(diagnostics,
						  DiagnosticName,
				          DIM1,
				          DIM2,
				          v,
				          vStartSegIndxs,
				          get_str_from_str_list(wks.GetPage().GetLongName(), "_", 0),
				          colPila.GetLongName(),
				          colSignal.GetLongName());
	
	Worksheet sheetOrigData = WksIdxByName(pg, "OrigData"),
			  sheetFitData = WksIdxByName(pg, "Fit"),
			  sheetParamsData = WksIdxByName(pg, "Parameters"),
			  sheetFiltData = WksIdxByName(pg, "Filtration"),
			  sheetDiffData = WksIdxByName(pg, "Diff");

    wks.ConnectTo(sheetOrigData);
    
	Dataset ds;
    for (i = 0; i < DIM1; ++i)
    {
    	/* Original Data sheet */
    	ds.Attach(sheetOrigData.Columns(0));
    	ds = vResP;
    	ds.Attach(sheetOrigData.Columns(i + 1));
		ds.SetSize(DIM2);
		vOriginalData.GetSubVector(ds, DIM2 * i, DIM2 * i + DIM2 - 1);
		
		/* Filtrated Data sheet */
    	if (filtS != 0)
		{
			ds.Attach(sheetFiltData.Columns(0));
			ds = vResP;
			ds.Attach(sheetFiltData.Columns(i + 1));
			ds.SetSize(DIM2);
			vFiltedData.GetSubVector(ds, DIM2 * i, DIM2 * i + DIM2 - 1);
		}
		
		if (diagnostics == 1)
		{
			ds.Attach(sheetDiffData.Columns(0));
			ds = vResP;
			ds.Attach(sheetDiffData.Columns(i + 1));
			ds.SetSize(DIM2);
			vDiffedData.GetSubVector(ds, DIM2 * i, DIM2 * i + DIM2 - 1);
		}
    	
    	/* Fit Data sheet */
    	ds.Attach(sheetFitData.Columns(0));
    	ds = vResP;
    	ds.Attach(sheetFitData.Columns(i + 1));
    	ds.SetSize(DIM2);
    	vResultData.GetSubVector(ds, DIM2 * i, DIM2 * i + DIM2 - 1);
    }
    
    /* Parameters */
    ds.Attach(sheetParamsData.Columns(0));
    ds = v;
    for (i = 1; i < Number_of_rezult_plasma_parametrs; ++i)
    {
    	ds.Attach(sheetParamsData.Columns(i));
    	ds.SetSize(DIM1);
    	for (j = i, k = 0; k < DIM1; j += Number_of_rezult_plasma_parametrs, ++k)
			ds[k] = vParameters[j];
    }
	
	set_active_layer(sheetParamsData);

	rezult_graphs(sheetOrigData.GetPage(), diagnostics);
	
	return 0;
}