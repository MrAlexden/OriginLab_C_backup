#include "AllinOneScript_rework.h"

int interf_prog(Column & colNTRF, int ch_dsc_ntrfrmtr, Worksheet & wks, double filtS, int vyr_or_not)
{
	if(!colNTRF) colNTRF = ColIdxByName(wks, s_ntrfrmtr);
	if(!colNTRF) return 666;

	//colNTRF.SetInternalData(FSI_DOUBLE);
	Dataset<double> dsDoub(colNTRF);
	Dataset<float> dsFlo(colNTRF);
	vector Signal;
	
	if(colNTRF.GetInternalData() == FSI_DOUBLE) 
		Signal = dsDoub;
	else
		Signal = dsFlo;
	
	if(Signal.GetUpperBound() <= 0) return 666;
	int i, k, prgrs = 1, size_sig = Signal.GetSize(), ind = 0;
	double lamb = 3.19*10^-3, me = 9.1*10^-31, eps = 8.85*10^-12, e = 1.6*10^-19, c = 3*10^8, tmax = (double)size_sig/(double)ch_dsc_ntrfrmtr;
	vector Phase;
	Phase.SetSize(size_sig);
	
	if(max(Signal) >= 10^5) return 555;
	
	progressBox prgbBox("Интерферометр " + wks.GetPage().GetLongName());
	prgbBox.SetRange(0, 9);
	
	if(!prgbBox.Set(prgrs++)) return 404;
	// Поиск точек пересечения исходным синусом нуля, запись в массив номеров этих точек
	k = 0;
	for(i = 0; i < size_sig - 1; i++)
	{
		if(Signal[i]<=0 && Signal[i+1]>=0)           // условие пересечения сигналом 0 снизу вверх
		{
			Phase[k]=(((-1)*Signal[i])/(Signal[i+1]-Signal[i]))+i;  // записываем точную координату пересечения 0 в отсчетах времени 
			k++;
		}
	}
	k = k - k%10000;
	
	if(!prgbBox.Set(prgrs++)) return 404;
	// Расчет длительности периодов в дискретах исходного сигнала
	for(i = 0; i < k - 1; i++, ind++)
		Phase[i] = Phase[i+1] - Phase[i];

	if(!prgbBox.Set(prgrs++)) return 404;
	// Пересчет нормированной к среднему длительности в градусы
	double Probe = 0;
	for(i = 0; i < k - 1; i++, ind++)
		Probe += Phase[i];
	
	Probe = Probe / (k-1);
	
	if(!prgbBox.Set(prgrs++)) return 404;
	for(i = 0; i < k - 1; i++, ind++)
		Phase[i] = Phase[i] / Probe;

	// вычисляем интегральный набег фазы (суммируя набеги за каждый из предшествующих периодов)
	double Phase0 = 0, Phase1 = 0;
	vector N, Time, N_posrednik;
	N.SetSize(k);
	N_posrednik.SetSize(k);
	Time.SetSize(k);
	
	if(!prgbBox.Set(prgrs++)) return 404;
	for(i = 1; i < k; i++, ind++)
	{
		Phase1 = Phase0 - ((Phase[i-1] - 1)*2*3.1416);
		N[i]=((4*3.1416*me*eps*Phase1*c^2)/(10^6*0.06*lamb*e^2));
		Phase0=Phase1;
	}
	
	if(!prgbBox.Set(prgrs++)) return 404;
	if(filtS != 0) 
    	ocmath_savitsky_golay(N, N_posrednik, N.GetSize(), N.GetSize()*filtS, -1, 5, 0, EDGEPAD_REFLECT);
	N=N_posrednik;

	if(!prgbBox.Set(prgrs++)) return 404;
	if(vyr_or_not == 1) Signal = N;
	else
	{
		int sizetodo = wks.Columns(0).GetUpperBound()+1, j = 0;
		if(k > sizetodo)
		{
			int rowstodel = k - sizetodo;
			int dIns = k/rowstodel;
			if(dIns == 1) dIns = 2;
			vector<int> removes(rowstodel);
			for(i = 0; j < rowstodel; i += dIns, j++)
			{
				if(i < k) 
					removes[j] = i;
				if(i >= k)
				{
					removes.SetSize(j);
					break;
				}
			}
			N.RemoveAt(removes);
		}
		Signal = N;
	}
	
	if(colNTRF.GetInternalData() == FSI_DOUBLE) 
		dsDoub = Signal;
	else
		dsFlo = Signal;
	
	int nR2;
	wks.GetBounds(0, 1, nR2, -1, false);
    wks.SetSize(nR2+1, -1);
	
	if(!prgbBox.Set(prgrs++)) return 404;
	
	return 0;
}