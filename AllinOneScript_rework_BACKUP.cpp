#include <Origin.h>
#include <..\originlab\NLFitSession.h>
#include <GetNBox.h>
#include <..\Originlab\graph_utils.h>// Needed for page_add_layer function

//////////////////////////////////// VARIABLES ////////////////////////////////////
int ch_dsc_p = 50000, ch_dsc_t = 200000, ch_pil = 50, r_z1 = 200, r_z2 = 10, r_s1 = 5000, r_s2 = 5000, Num_iter = 10, mnosj_sig = 30, r_cil = 1, 
	ch_dsc_ntrfrmtr = 1800000, koeff_rounding = 2500, NumPointsOfOriginalPila = 0, i_nach = 0, i_konec = 0;

double otsech_nachala = 0.05, otsech_konca = 0.05, par_A = -1, par_B = 1, par_C = 1, par_D = 10, koeff_width = 0.5, r0 = 0.006, d = 0.08, Ze = 1, 
	S1 = 3.141592*0.0005*0.005 + 3.141592*0.0005*0.0005/4, M_Ar = 6.6335209*10^-26 - 9.10938356*10^-31, M_He = 6.6464731*10^-27 - 9.10938356*10^-31, razbros_po_pile = 0.25, 
	S2 = 3.141592*0.002*0.003 + 3.141592*0.002*0.002/4, c_A = 5, c_B = 2, c_C = 80, c_D = 12, c_E = 50, m_A = 500, m_B = 1, m_C = 1, m_D = 1, m_E = 0, chast_to_fitln = 0.5, 
	AverageSignal = 0, st_time_end_time[2] = {-1, -1};
	
string s_pz = "Пила Зонда_ЦАП", s_tz = "ТокЗонда", s_tz2 = "Ток Зонда 2", s_ps = "Пила Сетки_ЦАП", s_tst = "Ток Сеточный Торец", s_tssh = "Ток Сеточный Штанга",
	svch = "СВЧ", vch = "ВЧ", icr = "ИЦР", s_pc = "Пила цилиндр", s_tc = "ТокЦилиндрический", s_ntrfrmtr = "Интерферометр",
	s_pm = "ПилаМагнитный", s_tm = "ТокМагнитный";
//////////////////////////////////// VARIABLES ////////////////////////////////////

void zond()
{
	Worksheet wks = Project.ActiveLayer(), wks_slave;
    if(!wks) 
    {
    	MessageBox(GetWindow(), "wks!");
    	return;
    }
    
    if(wks.GetPage().GetLongName().Match("Зонд*")) 
    {
    	//pereschet_impulse_zond(wks);
    	MessageBox(GetWindow(), "Не тот лист! Выберите лист с результатами эксперимента");
    	return;
    }
    
	if(wks.GetPage().GetLongName().Match("Зонд*") || wks.GetPage().GetLongName().Match("Сетка*") 
		|| wks.GetPage().GetLongName().Match("Цилиндр*") || wks.GetPage().GetLongName().Match("Магнит*"))
	{
		MessageBox(GetWindow(), "Не тот лист! Выберите лист с результатами эксперимента");
		return;
	}

    vector<string> vBooksNames, vErrsBooksNames;
	foreach(WorksheetPage pg in Project.WorksheetPages)
	{
		int BeginVal;
		is_str_numeric_integer(get_str_from_str_list(pg.GetLongName(), "_", 0), &BeginVal);
		if(BeginVal>10000) vBooksNames.Add(pg.GetLongName());
	}
	
    GETN_BOX( treeTest );
    GETN_NUM(decayT2, "Частота пилы, Гц", ch_pil)
    
    GETN_BEGIN_BRANCH(Start, "Fitting Options")
    GETN_MULTI_COLS_BRANCH(3, DYNA_MULTI_COLS_COMAPCT) GETN_OPTION_BRANCH(GETNBRANCH_OPEN)
		
		GETN_CHECK(decayT3,"Обработка данных c зонда 1", true)
		GETN_NUM(decayT5, "Площадь поверхности, м^2", S1) GETN_EDITOR_SIZE_USER_OPTION("#12")
		GETN_NUM(decayT1, "Сопротивление на зонде, Ом", r_z1) GETN_EDITOR_SIZE_USER_OPTION("#5")
		GETN_CHECK(decayT4,"Обработка данных с зонда 2", false)
		GETN_NUM(decayT6, "Площадь поверхности, м^2", S2) GETN_EDITOR_SIZE_USER_OPTION("#12")
		GETN_NUM(decayT2, "Сопротивление на зонде, Ом", r_z2) GETN_EDITOR_SIZE_USER_OPTION("#5")
		
    GETN_END_BRANCH(Start)
    
    GETN_BEGIN_BRANCH(Details2, "Задать частоты дискретизации(по умолчанию автоопределение)") GETN_CHECKBOX_BRANCH(0)
    //GETN_OPTION_BRANCH(GETNBRANCH_CLOSED)
	GETN_NUM(decayT8, "Частота дискретизации пилы, Гц", ch_dsc_p)
    GETN_NUM(decayT9, "Частота дискретизации зонда, Гц", ch_dsc_t)
    GETN_END_BRANCH(Details2)
    
    GETN_BEGIN_BRANCH(Details3, "Задать начало и конец обработки(по умолчанию автоопределение)") GETN_CHECKBOX_BRANCH(0)
    //GETN_OPTION_BRANCH(GETNBRANCH_CLOSED)
    GETN_SLIDEREDIT(decayT1, "Время начала, с", 0.3, "0.0|5.0|50") 
    GETN_SLIDEREDIT(decayT2, "Время конца, с", 1.5, "0.0|5.0|50") 
    GETN_END_BRANCH(Details3)
    
    GETN_SLIDEREDIT(decayT14, "Часть точек, которые будут отсечены от начала", otsech_nachala, "0.01|0.49|48")
    GETN_SLIDEREDIT(decayT15, "Часть точек, которые будут отсечены от конца", otsech_konca, "0.01|0.49|48")
    
    GETN_SLIDEREDIT(decayT12, "Часть данных от начала которые будут линейно аппроксимированны", chast_to_fitln, "0.00|0.99|99")
    GETN_RADIO_INDEX_EX(decayT11, "Тип импульса", 2, "СВЧ|ВЧ|ИЦР")
    GETN_RADIO_INDEX_EX(decayT13, "Рабочее тело", 0, "He|Ar")
    
    GETN_BEGIN_BRANCH(Details1, "Фильтрация данных") GETN_CHECKBOX_BRANCH(0)
    GETN_SLIDEREDIT(decayT14, "Часть точек фильтрации", 0.1, "0.01|0.99|98") 
    GETN_END_BRANCH(Details1)
    
    GETN_BEGIN_BRANCH(Details, "Нужно обработать несколько книг") GETN_CHECKBOX_BRANCH(0)
    for(int i_1 = 0; i_1<vBooksNames.GetSize(); i_1++) 
    {
		GETN_CHECK(NodeName, vBooksNames[i_1], true)
    }
    GETN_END_BRANCH(Details)
    
    if(0==GetNBox( treeTest, "Обработка данных с зонда" )) return;
    
    double chastota_discret_pily, chastota_discret_zonda, Chast_to_Filtrate;
    int Obr_Z1, Obr_Z2, SVCH_or_VCH, He_or_Ar, True_or_False, Error = -1, UseSameName_or_Not_PILA, UseSameName_or_Not_ZONDA, UseSameName_or_Not_ZONDA_2, z_s_c = 0;
    vector vecErrs;
    TreeNode tn1;
    
    //////////////////////////// СЧИТЫВАЕМ ДАННЫЕ ИЗ ДИАЛОГА ////////////////////////////
    treeTest.GetNode("Start").GetNode("decayT3").GetValue(Obr_Z1);
	treeTest.GetNode("Start").GetNode("decayT4").GetValue(Obr_Z2);
	treeTest.GetNode("Start").GetNode("decayT5").GetValue(S1);
	treeTest.GetNode("Start").GetNode("decayT6").GetValue(S2);
	treeTest.GetNode("Start").GetNode("decayT1").GetValue(r_z1);
	treeTest.GetNode("Start").GetNode("decayT2").GetValue(r_z2);
	treeTest.GetNode("decayT2").GetValue(ch_pil);
	treeTest.GetNode("decayT11").GetValue(SVCH_or_VCH);
	treeTest.GetNode("decayT13").GetValue(He_or_Ar);
	treeTest.GetNode("decayT12").GetValue(chast_to_fitln);
	
	treeTest.GetNode("decayT14").GetValue(otsech_nachala);	
	treeTest.GetNode("decayT15").GetValue(otsech_konca);
	otsech_nachala = (otsech_nachala == 0) ? 0.01 : otsech_nachala;
	otsech_konca = (otsech_konca == 0) ? 0.01 : otsech_konca;
	
	if(treeTest.Details1.Use)
	{
		treeTest.GetNode("Details1").GetNode("decayT14").GetValue(Chast_to_Filtrate);
		Chast_to_Filtrate = Chast_to_Filtrate/2;
	}
	else Chast_to_Filtrate = 0;
	if(treeTest.Details3.Use)
	{
		treeTest.GetNode("Details3").GetNode("decayT1").GetValue(st_time_end_time[0]);
		treeTest.GetNode("Details3").GetNode("decayT2").GetValue(st_time_end_time[1]);
		if(st_time_end_time[0] >= st_time_end_time[1])
		{
			MessageBox(GetWindow(), "Ошибка. Исправьте время начала и конца обработки. Вы выбрали от " + st_time_end_time[0] + " до " + st_time_end_time[1]);
			return;
		}
	}
	//////////////////////////// СЧИТЫВАЕМ ДАННЫЕ ИЗ ДИАЛОГА ////////////////////////////
   	
	string ReportString_ZONDA, ReportString_ZONDA_2, ReportString_PILA;
   	Column colPILA = ColIdxByName(wks, s_pz);
	if(!colPILA)
	{
		GETN_BOX( treeTest1 );
		GETN_STR(STR, "!Не найден столбец с именем " + "\"" + s_pz + "\"" + "!", "") GETN_INFO
		GETN_INTERACTIVE(interactive, "Выберите столбец с нужными данными", "[Book1]Sheet1!B")
		GETN_OPTION_INTERACTIVE_CONTROL(ICOPT_RESTRICT_TO_ONE_DATA) 
		if(treeTest.Details.Use)
		{
			GETN_RADIO_INDEX_EX(decayT1, "Использовать колонку с таким же именем для расчета остальных книг", 0, "Да|Нет")
		}
		if(0==GetNBox( treeTest1, "Ошибка!" )) return;
		
		treeTest1.GetNode("interactive").GetValue(ReportString_PILA);
		treeTest1.GetNode("decayT1").GetValue(UseSameName_or_Not_PILA);
		ReportString_PILA = get_str_from_str_list(ReportString_PILA, "\"", 1);
		colPILA = ColIdxByName(wks, ReportString_PILA);
		if(!colPILA)
		{
			MessageBox(GetWindow(), "Ошибка. Возможно, выбранная вами колонка пустая");
			return;
		}
		colPILA.SetLongName(s_pz);
		UseSameName_or_Not_PILA = UseSameName_or_Not_PILA + 2;
	}
	Column colTOK_ZONDA, colTOK_ZONDA_2; 
	if(Obr_Z1 == 1)
	{
		colTOK_ZONDA = ColIdxByName(wks, s_tz);
		if(!colTOK_ZONDA)
		{
			GETN_BOX( treeTest1 );
			GETN_STR(STR, "!Не найден столбец с именем " + "\"" + s_tz + "\"" + "!", "") GETN_INFO
			GETN_INTERACTIVE(interactive, "Выберите столбец с нужными данными", "[Book1]Sheet1!B")
			GETN_OPTION_INTERACTIVE_CONTROL(ICOPT_RESTRICT_TO_ONE_DATA) 
			if(treeTest.Details.Use) 
			{
				GETN_RADIO_INDEX_EX(decayT1, "Использовать колонку с таким же именем для расчета остальных книг", 0, "Да|Нет")
			}
			if(0==GetNBox( treeTest1, "Ошибка!" )) return;
			
			treeTest1.GetNode("interactive").GetValue(ReportString_ZONDA);
			treeTest1.GetNode("decayT1").GetValue(UseSameName_or_Not_ZONDA);
			ReportString_ZONDA = get_str_from_str_list(ReportString_ZONDA, "\"", 1);
			colTOK_ZONDA = ColIdxByName(wks, ReportString_ZONDA);
			if(!colTOK_ZONDA)
			{
				MessageBox(GetWindow(), "Ошибка. Возможно, выбранная вами колонка пустая");
				return;
			}
			colTOK_ZONDA.SetLongName(s_tz);
			UseSameName_or_Not_ZONDA = UseSameName_or_Not_ZONDA + 2;
		}
	}
	if(Obr_Z2 == 1)
	{
		colTOK_ZONDA_2 = ColIdxByName(wks, s_tz2);
		if(!colTOK_ZONDA_2)
		{
			GETN_BOX( treeTest1 );
			GETN_STR(STR, "!Не найден столбец с именем " + "\"" + s_tz2 + "\"" + "!", "") GETN_INFO
			GETN_INTERACTIVE(interactive, "Выберите столбец с нужными данными", "[Book1]Sheet1!B")
			GETN_OPTION_INTERACTIVE_CONTROL(ICOPT_RESTRICT_TO_ONE_DATA) 
			if(treeTest.Details.Use) 
			{
				GETN_RADIO_INDEX_EX(decayT1, "Использовать колонку с таким же именем для расчета остальных книг", 0, "Да|Нет")
			}
			if(0==GetNBox( treeTest1, "Ошибка!" )) return;
			
			treeTest1.GetNode("interactive").GetValue(ReportString_ZONDA_2);
			treeTest1.GetNode("decayT1").GetValue(UseSameName_or_Not_ZONDA_2);
			ReportString_ZONDA_2 = get_str_from_str_list(ReportString_ZONDA_2, "\"", 1);
			colTOK_ZONDA_2 = ColIdxByName(wks, ReportString_ZONDA_2);
			if(!colTOK_ZONDA_2)
			{
				MessageBox(GetWindow(), "Ошибка. Возможно, выбранная вами колонка пустая");
				return;
			}
			colTOK_ZONDA_2.SetLongName(s_tz2);
			UseSameName_or_Not_ZONDA_2 = UseSameName_or_Not_ZONDA_2 + 2;
		}
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
				Column colTOK_ZONDA_slave, colTOK_ZONDA_2_slave, colPILA_slave;
				Page pg1 = Project.FindPage(vBooksNames[i_2]);
				if(pg1.Layers(0) != wks) 
				{
					wks_slave = pg1.Layers(0);
					if(UseSameName_or_Not_ZONDA == 2) 
					{
						colTOK_ZONDA_slave = ColIdxByName(wks_slave, ReportString_ZONDA);
						if(!colTOK_ZONDA_slave)
						{
							vecErrs.Add(222);
							vErrsBooksNames.Add("Обработка зонда "+vBooksNames[i_2]+": ");
							continue;
						}
					}
					if(UseSameName_or_Not_ZONDA_2 == 2) 
					{
						colTOK_ZONDA_2_slave = ColIdxByName(wks_slave, ReportString_ZONDA_2);
						if(!colTOK_ZONDA_2_slave)
						{
							vecErrs.Add(222);
							vErrsBooksNames.Add("Обработка зонда "+vBooksNames[i_2]+": ");
							continue;
						}
					}
					if(UseSameName_or_Not_PILA == 2) 
					{
						colPILA_slave = ColIdxByName(wks_slave, ReportString_PILA);
						if(!colPILA_slave)
						{
							vecErrs.Add(333);
							vErrsBooksNames.Add("Пила зонда "+vBooksNames[i_2]+": ");
							continue;
						}
					}
				}
				else
				{
					wks_slave = wks;
					colTOK_ZONDA_slave = colTOK_ZONDA;
					colTOK_ZONDA_2_slave = colTOK_ZONDA_2;
					colPILA_slave = colPILA;
				}
				colTOK_ZONDA_slave.SetLongName(s_tz);
				colTOK_ZONDA_2_slave.SetLongName(s_tz2);
				colPILA_slave.SetLongName(s_pz);

				if(Obr_Z1 == 1 && Error < 0) 
				{
					if(!treeTest.Details2.Use)
					{
						Error = auto_chdsc(chastota_discret_pily, chastota_discret_zonda, colPILA_slave, colTOK_ZONDA_slave, ch_pil, z_s_c, wks_slave, 0);
						if( Error != -1 ) 
						{
							vecErrs.Add(Error);
							if(Error != 333) vErrsBooksNames.Add("Обработка зонда1 "+vBooksNames[i_2]+": ");
							else vErrsBooksNames.Add("Пила зонда "+vBooksNames[i_2]+": ");
						}
						Error = -1;
					}
					else
					{
						treeTest.GetNode("Details2").GetNode("decayT8").GetValue(chastota_discret_pily);
						treeTest.GetNode("Details2").GetNode("decayT9").GetValue(chastota_discret_zonda);
					}
					
					Error = zond_obrab(r_z1, ch_pil, chastota_discret_pily, chastota_discret_zonda, SVCH_or_VCH, He_or_Ar, 
										wks_slave, colPILA_slave, colTOK_ZONDA_slave, Chast_to_Filtrate, z_s_c, 0, S1, c_A, c_B, c_C, c_D, c_E);
					if(Error == 404) return;
					if(Error != -1) 
					{
						vecErrs.Add(Error);
						if(Error != 333) vErrsBooksNames.Add("Обработка зонда1 "+vBooksNames[i_2]+": ");
						else vErrsBooksNames.Add("Пила зонда "+vBooksNames[i_2]+": ");
					}
					Error = -1;
				}
				if(Obr_Z2 == 1 && Error < 0) 
				{
					if(!treeTest.Details2.Use)
					{
						Error = auto_chdsc(chastota_discret_pily, chastota_discret_zonda, colPILA_slave, colTOK_ZONDA_2_slave, ch_pil, z_s_c, wks_slave, 1);
						if( Error != -1 ) 
						{
							vecErrs.Add(Error);
							if(Error != 333) vErrsBooksNames.Add("Обработка зонда1 "+vBooksNames[i_2]+": ");
							else vErrsBooksNames.Add("Пила зонда "+vBooksNames[i_2]+": ");
						}
						Error = -1;
					}
					else
					{
						treeTest.GetNode("Details2").GetNode("decayT8").GetValue(chastota_discret_pily);
						treeTest.GetNode("Details2").GetNode("decayT9").GetValue(chastota_discret_zonda);
					}
					
					Error = zond_obrab(r_z2, ch_pil, chastota_discret_pily, chastota_discret_zonda, SVCH_or_VCH, He_or_Ar, 
										wks_slave, colPILA_slave, colTOK_ZONDA_2_slave, Chast_to_Filtrate, z_s_c, 1, S2, c_A, c_B, c_C, c_D, c_E);
					if(Error == 404) return;
					if(Error != -1) 
					{
						vecErrs.Add(Error);
						if(Error != 333) vErrsBooksNames.Add("Обработка зонда2 "+vBooksNames[i_2]+": ");
						else vErrsBooksNames.Add("Пила зонда "+vBooksNames[i_2]+": ");
					}
					Error = -1;
				}
			}
		}
	}
	else 
	{
		if(Obr_Z1 == 1) 
		{
			if(!treeTest.Details2.Use)
			{
				Error = auto_chdsc(chastota_discret_pily, chastota_discret_zonda, colPILA, colTOK_ZONDA, ch_pil, z_s_c, wks, 0);
				if( Error != -1 )
				{
					MessageBox(GetWindow(), FuncErrors(Error));
					return;
				}
				Error = -1;
			}
			else
			{
				treeTest.GetNode("Details2").GetNode("decayT8").GetValue(chastota_discret_pily);
				treeTest.GetNode("Details2").GetNode("decayT9").GetValue(chastota_discret_zonda);
			}
			
			Error = zond_obrab(r_z1, ch_pil, chastota_discret_pily, chastota_discret_zonda, SVCH_or_VCH, He_or_Ar, 
								wks, colPILA, colTOK_ZONDA, Chast_to_Filtrate, z_s_c, 0, S1, c_A, c_B, c_C, c_D, c_E);
			if(Error == 404) return;
			if( Error != -1 ) 
			{
				vecErrs.Add(Error);
				if(Error != 333) vErrsBooksNames.Add("Зонд1 "+get_str_from_str_list(wks.GetPage().GetLongName(), "_", 0)+": ");
				else vErrsBooksNames.Add("Пила Зонда "+get_str_from_str_list(wks.GetPage().GetLongName(), "_", 0)+": ");
			}
			Error = -1;
		}
		if(Obr_Z2 == 1) 
		{
			if(!treeTest.Details2.Use)
			{
				Error = auto_chdsc(chastota_discret_pily, chastota_discret_zonda, colPILA, colTOK_ZONDA_2, ch_pil, z_s_c, wks, 1);
				if( Error != -1 )
				{
					MessageBox(GetWindow(), FuncErrors(Error));
					return;
				}
				Error = -1;
			}
			else
			{
				treeTest.GetNode("Details2").GetNode("decayT8").GetValue(chastota_discret_pily);
				treeTest.GetNode("Details2").GetNode("decayT9").GetValue(chastota_discret_zonda);
			}
			
			Error = zond_obrab(r_z2, ch_pil, chastota_discret_pily, chastota_discret_zonda, SVCH_or_VCH, He_or_Ar,
								wks, colPILA, colTOK_ZONDA_2, Chast_to_Filtrate, z_s_c, 1, S2, c_A, c_B, c_C, c_D, c_E);
			if(Error == 404) return;
			if( Error != -1 ) 
			{
				vecErrs.Add(Error);
				if(Error != 333) vErrsBooksNames.Add("Зонд2 "+get_str_from_str_list(wks.GetPage().GetLongName(), "_", 0)+": ");
				else vErrsBooksNames.Add("Пила Зонда "+get_str_from_str_list(wks.GetPage().GetLongName(), "_", 0)+": ");
			}
			Error = -1;
		}
	}
	if(vecErrs.GetSize()!=0)
	{
		string str = "Были ошибки в следующих книгах: \n";
		for(int aboba=0; aboba<vecErrs.GetSize(); aboba++) str = str + vErrsBooksNames[aboba] + FuncErrors(vecErrs[aboba]) + "\n";
		MessageBox(GetWindow(), str);
	}
	
	////////////////// возвращяю эту хрень к изначальным значениям иначе он ее запоминает и ломается //////////////////
	st_time_end_time[0] = -1;
	st_time_end_time[1] = -1;
	////////////////// возвращяю эту хрень к изначальным значениям иначе он ее запоминает и ломается //////////////////
}

int zond_obrab(double resistance, double chastota_pily, double chastota_discret_pily, double chastota_discret_zonda, int SVCH_or_VCH, 
				int He_or_Ar, Worksheet wks, Column colPILA, Column colTOK_ZONDA, double Chast_to_Filtrate, int z_s_c, int diagnostic_num, double S,
				double c_A, double c_B, double c_C, double c_D, double c_E)
{
	time_t start_time, finish_time;
	time(&start_time);
	
	switch (z_s_c)
	{
		case 0:
			if(!colPILA) colPILA = ColIdxByName(wks, s_pz);	
			if(diagnostic_num == 0 && !colTOK_ZONDA) colTOK_ZONDA = ColIdxByName(wks, s_tz);
			if(diagnostic_num == 1 && !colTOK_ZONDA) colTOK_ZONDA = ColIdxByName(wks, s_tz2);
			break;
		case 2:
			if(diagnostic_num == 0 && !colPILA) colPILA = ColIdxByName(wks, s_pc);
			if(diagnostic_num == 1 && !colPILA) colPILA = ColIdxByName(wks, s_pm);
			if(diagnostic_num == 0 && !colTOK_ZONDA) colTOK_ZONDA = ColIdxByName(wks, s_tc);
			if(diagnostic_num == 1 && !colTOK_ZONDA) colTOK_ZONDA = ColIdxByName(wks, s_tm);
			break;
	}

	if(!colTOK_ZONDA) return 222;
	if(!colPILA) return 333;
   	vector vecPILA_ZONDA(colPILA), vecTOK_ZONDA(colTOK_ZONDA), vec_sokrashPILA, vec_Otr_Start_Indxs(0);
    int kolvo_otrezkov=0;
 
	if(z_s_c==0) vecPILA_ZONDA = vecPILA_ZONDA * mnosj_sig;
	else vecPILA_ZONDA = vecPILA_ZONDA * (c_A*c_B*(c_C/c_D)) - c_E;
   		
	if( find_otrezki_toka_i_obrab_pily(vec_Otr_Start_Indxs, vecTOK_ZONDA, chastota_discret_zonda, chastota_pily, vecPILA_ZONDA, chastota_discret_pily, vec_sokrashPILA) == 0 ) return 444;
	kolvo_otrezkov = vec_Otr_Start_Indxs.GetSize()+1;

	vector vecOBRAB_TOKA, vecTIME_synth, vec_X_stepennoi, vecFiltrate, vecA, vecB, vecC, vecD;
	vecTIME_synth.SetSize(kolvo_otrezkov-1);
	vecA.SetSize(kolvo_otrezkov-1);
	vecB.SetSize(kolvo_otrezkov-1);
	vecC.SetSize(kolvo_otrezkov-1);
	vecD.SetSize(kolvo_otrezkov-1);
	vecOBRAB_TOKA.SetSize(vec_sokrashPILA.GetSize()*kolvo_otrezkov);
	vecFiltrate.SetSize(vec_sokrashPILA.GetSize()*kolvo_otrezkov);
	vec_X_stepennoi = vec_sokrashPILA;
	
	/* переворачиваем ток чтобы смотрел вверх(если нужно), и делим на сопротивление */
	if(z_s_c != 0) turn_tok_upwards(vecTOK_ZONDA, resistance, -1);
	else
	{
		vecTOK_ZONDA /= -resistance;
		vec_Otr_Start_Indxs *= 1.001; //КОСТЫЛЬ
	}
	
	time_t start_time1, finish_time1;
	time(&start_time1);
	
	string s_progress, pgOrigName = get_str_from_str_list(wks.GetPage().GetLongName(), "_", 0);
	switch (z_s_c)
	{
		case 0:
			if(diagnostic_num == 0) s_progress = "Обработка зонда1 ";
			if(diagnostic_num == 1) s_progress = "Обработка зонда2 ";
			break;
		case 2:
			if(diagnostic_num == 0) s_progress = "Обработка цилиндрического ";
			if(diagnostic_num == 1) s_progress = "Обработка магнитного ";
			break;
	}
	
	progressBox prgbBox(s_progress + wks.GetPage().GetLongName());
	prgbBox.SetRange(0, kolvo_otrezkov-1);
	
	int indFilt = 0, i = 0, lolo = 0;
	for(i=0, lolo = 0; i<kolvo_otrezkov-1; i++)
	{
		if(!prgbBox.Set(i)) return 404;
		
		vecTIME_synth[i] = vec_Otr_Start_Indxs[i]*(1/chastota_discret_zonda);	
		
		make_one_segment(z_s_c, vec_Otr_Start_Indxs, vecTOK_ZONDA, chastota_discret_zonda, chastota_pily, vec_sokrashPILA, Chast_to_Filtrate, vecFiltrate, 
							vecA, vecB, vecC, vecD, vecOBRAB_TOKA, -1, -1, lolo, -1, i, indFilt);
	}
	
	time(&finish_time1);
	printf("Run time fitting %d secs\n", finish_time1 - start_time1);
	
	vector vecDens;
	if(z_s_c==0) dens(true, 0, vecDens, vecA, vecB, vecC, vecD, vec_sokrashPILA, He_or_Ar, S);
	
	Worksheet wks_test, wks_test_1, wks_test_2, wks_test_3;
	Page pg;
	wks_test.Create("Origin");
	pg = wks_test.GetPage();
	wks.ConnectTo(wks_test);
	
	create_wks_z_s_c(z_s_c, pg, SVCH_or_VCH, He_or_Ar, kolvo_otrezkov, vec_sokrashPILA, Chast_to_Filtrate, resistance, chastota_discret_pily, chastota_discret_zonda, 
					vec_Otr_Start_Indxs, chastota_pily, diagnostic_num, -1, -1, pgOrigName);
	
	wks_test = WksIdxByName(pg, "OrigData");
    wks_test_1 = WksIdxByName(pg, "Fit");
    wks_test_2 = WksIdxByName(pg, "Parameters");
    wks_test_3 = WksIdxByName(pg, "Filtration");
	
	Dataset dsTIME_synth(wks_test_2.Columns(0)), dsA(wks_test_2.Columns(1)), dsB(wks_test_2.Columns(2)), dsC(wks_test_2.Columns(3)), dsD(wks_test_2.Columns(4)), dsDens(wks_test_2.Columns(5)), ds_STOLBOV(wks_test.Columns(0));
	
	dsTIME_synth=vecTIME_synth;
	dsA=vecA;
	dsB=vecB;
	dsC=vecC;
	dsD=vecD;
	if(z_s_c==0)
		dsDens=vecDens;
	ds_STOLBOV=vec_sokrashPILA;
	ds_STOLBOV.Attach(wks_test_1.Columns(0));
	ds_STOLBOV=vec_sokrashPILA;
	if(Chast_to_Filtrate != 0)
	{
		ds_STOLBOV.Attach(wks_test_3.Columns(0));
		ds_STOLBOV=vec_sokrashPILA;
	}

	indFilt=0;
	for(int i_2=1, k=0; i_2<kolvo_otrezkov; i_2++)
	{
		ds_STOLBOV.Attach(wks_test.Columns(i_2));
		ds_STOLBOV.SetSize(vec_sokrashPILA.GetSize());
		
		for(int i=0, k_1 = vec_Otr_Start_Indxs[i_2-1] + (chastota_discret_zonda/chastota_pily)*otsech_nachala; i < vec_sokrashPILA.GetSize(); i++, k_1++)
		{
			ds_STOLBOV[i]=vecTOK_ZONDA[k_1];
		}
		
		if(Chast_to_Filtrate != 0)
		{
			ds_STOLBOV.Attach(wks_test_3.Columns(i_2));
			ds_STOLBOV.SetSize(vec_sokrashPILA.GetSize());
			for(int i=0; i<vec_sokrashPILA.GetSize(); i++, indFilt++)
			{
				ds_STOLBOV[i]=vecFiltrate[indFilt];
			}
		}
		
		ds_STOLBOV.Attach(wks_test_1.Columns(i_2));
		ds_STOLBOV.SetSize(vec_sokrashPILA.GetSize());
		for(int j=0; j<vec_sokrashPILA.GetSize(); j++, k++)
		{
			ds_STOLBOV[j]=vecOBRAB_TOKA[k]
		}
	}
	wks_test_3.SetIndex(1);
	
	set_active_layer(wks_test_2);

	rezult_graphs(wks_test.GetPage(), z_s_c);
	
	time(&finish_time);
	printf("Run time all prog %d secs\n", finish_time - start_time);
	
	return -1;
}

void make_one_segment(int z_s_c, vector &vec_Otr_Start_Indxs, vector &vecTOK, double chastota_discret_toka, double chastota_pily, vector &vec_sokrashPILA, 
						double Chast_to_Filtrate, vector &vecFiltrate, vector &vecA, vector &vecB, vector &vecC, vector &vecD, 
						vector &vecOBRAB_TOKA, double Chast_to_Filtrate_Diff, int Poly_Order, int &lolo, int &lolo_1, int &i_1, int &indFilt)
{	
	switch (z_s_c)
	{
		case 0:
			
			int i, k;
			vector vec_Y_stepennoi, vec_Y_stepennoi_posrednik, vParams(4), vec_X_stepennoi_posrednik, vec_X_stepennoi; 
			vec_X_stepennoi = vec_sokrashPILA;
			vec_Y_stepennoi.SetSize(vec_sokrashPILA.GetSize());	
			vec_Y_stepennoi_posrednik.SetSize(vec_sokrashPILA.GetSize());	
			
			for(i = 0, k = vec_Otr_Start_Indxs[i_1] + (chastota_discret_toka/chastota_pily)*otsech_nachala; i < vec_sokrashPILA.GetSize(); i++, k++) 
			{	
				vec_Y_stepennoi[i] = vecTOK[k];
			}
			
			if(Chast_to_Filtrate != 0)
			{
				ocmath_savitsky_golay(vec_Y_stepennoi, vec_Y_stepennoi_posrednik, vec_Y_stepennoi.GetSize(), vec_Y_stepennoi.GetSize()*Chast_to_Filtrate, -1, 5, 0, EDGEPAD_REFLECT);
				vec_Y_stepennoi = vec_Y_stepennoi_posrednik;
				for(i=0; i<vec_Y_stepennoi.GetSize(); i++, indFilt++) 
					vecFiltrate[indFilt] = vec_Y_stepennoi[i];
			}
		
			vParams[0] = vec_Y_stepennoi[0];
			vParams[1] = par_B;
			
			vec_X_stepennoi_posrednik.SetSize(vec_Y_stepennoi.GetSize()*0.5);
			vec_Y_stepennoi_posrednik.SetSize(vec_Y_stepennoi.GetSize()*0.5);
			for(i = vec_Y_stepennoi.GetSize()*0.5, k = 0; i < vec_Y_stepennoi.GetSize()-1; i++, k++)
			{
				vec_Y_stepennoi_posrednik[k] = vec_Y_stepennoi[i];
				vec_X_stepennoi_posrednik[k] = vec_X_stepennoi[i];
			}
			
			NLFitSession FitSession;
			FitSession.SetFunction("Stepennaya_BAX");
			FitSession.SetMaxNumIter(Num_iter);
			FitSession.SetData(vec_Y_stepennoi_posrednik, vec_X_stepennoi_posrednik);
			
			if(chast_to_fitln != 0) koeff_fit_linear_stepennaya(vParams, vec_Y_stepennoi, vec_X_stepennoi);
			
			vParams[2] = -vParams[0];
			vParams[3] = par_D;
			
			FitSession.SetParamValues( vParams );
		
			if(chast_to_fitln != 0)
			{
				FitSession.SetParamFix(0, true);
				FitSession.SetParamFix(1, true);
			}
			
			int     nOutcome;
			FitSession.Fit(&nOutcome, false, false);
			
			vector vParamValues, vErrors;
			FitSession.GetFitResultsParams(vParamValues, vErrors);
			
			vecA[i_1]=vParamValues[0];
			vecB[i_1]=vParamValues[1];
			if(vParamValues[2] > 0) vecC[i_1] = vParamValues[2];
			else vecC[i_1] = -vParamValues[0]/4;
			if((vParamValues[3]<0 || vParamValues[3]>100) && i_1!=0) vecD[i_1] = vecD[i_1-1];
			else vecD[i_1] = vParamValues[3];
			
			for (i = 0; i < vec_sokrashPILA.GetSize(); i++, lolo++)
			{
				vecOBRAB_TOKA[lolo] = fx(vec_X_stepennoi[i],vParamValues[0],vParamValues[1],vParamValues[2],vParamValues[3]);
			}
			
			FitSession.SetParamFix(0, false);
			FitSession.SetParamFix(1, false);
			
			break;
			
		case 1:
			
			int i, k;
			vector vec_Y_stepennoi, vec_Y_stepennoi_posrednik, vec_X_stepennoi; 
			vec_X_stepennoi = vec_sokrashPILA;
			vec_Y_stepennoi.SetSize(vec_sokrashPILA.GetSize());
			vec_Y_stepennoi_posrednik.SetSize(vec_sokrashPILA.GetSize());

			for(i = 0, k = vec_Otr_Start_Indxs[i_1] + (chastota_discret_toka/chastota_pily)*otsech_nachala; i < vec_sokrashPILA.GetSize(); i++, k++) 
			{	
				vec_Y_stepennoi[i] = vecTOK[k];
			}
			
			ocmath_savitsky_golay(vec_Y_stepennoi, vec_Y_stepennoi_posrednik, vec_sokrashPILA.GetSize(), vec_sokrashPILA.GetSize()*Chast_to_Filtrate, -1, 5, 0, EDGEPAD_REFLECT);
			vec_Y_stepennoi = vec_Y_stepennoi_posrednik;
			
			for (i = 0; i < vec_sokrashPILA.GetSize(); i++, lolo++)
			{
				vecOBRAB_TOKA[lolo] = vec_Y_stepennoi[i];
			}
			if(Chast_to_Filtrate_Diff == 0) ocmath_derivative( vec_X_stepennoi, vec_Y_stepennoi, vec_sokrashPILA.GetSize());
			else ocmath_savitsky_golay(vec_Y_stepennoi_posrednik, vec_Y_stepennoi, vec_sokrashPILA.GetSize(), vec_sokrashPILA.GetSize()*Chast_to_Filtrate_Diff, -1, Poly_Order, 1, EDGEPAD_REFLECT);
			
			vec_Y_stepennoi = -vec_Y_stepennoi;
			
			for (i = 0; i < vec_sokrashPILA.GetSize(); i++, lolo_1++)
			{
				vecFiltrate[lolo_1] = vec_Y_stepennoi[i];
			}
			
			vector v_whm_w_peak(10);
			/*
			вектор v_whm_w_peak:
				whm
				-[0] ширина на полувысоте
				-[1] Y на полувысоте
				-[2] X на полувысоте слева
				-[3] X на полувысоте справа
				W
				-[4] значение коэффициента W
				-[5] Y при W
				-[6] X при W слева
				-[7] X при W справа
				Peak
				-[8] Y пика
				-[9] X пика
			*/
			v_whm_w_peak = find_whm_Wparam_LRmidindxs(vec_X_stepennoi, vec_Y_stepennoi, 0, 0, -1);
			
			vecA[i_1] = v_whm_w_peak[8]; //Max Value
			vecB[i_1] = v_whm_w_peak[0]; //Temp
			vecC[i_1] = v_whm_w_peak[9]; //Peak Voltage
			if(Ze*r0*v_whm_w_peak[9]/2*d < 1000 || i_1 == 0) 
				vecD[i_1] = Ze*r0*v_whm_w_peak[9]/2*d; //Energy
			else if(i_1 != 0)
				vecD[i_1] = vecD[i_1-1]; //Energy
			
			break;
			
		case 2:
			
			int i, k, up_down = 0;
			vector vec_Y_stepennoi, vec_Y_stepennoi_posrednik, vec_X_stepennoi; 
			vec_X_stepennoi = vec_sokrashPILA;
			vec_Y_stepennoi.SetSize(vec_sokrashPILA.GetSize());	
			vec_Y_stepennoi_posrednik.SetSize(vec_sokrashPILA.GetSize());	
			
			for(i = 0, k = vec_Otr_Start_Indxs[i_1] + (chastota_discret_toka/chastota_pily)*otsech_nachala; i < vec_sokrashPILA.GetSize(); i++, k++) 
			{	
				vec_Y_stepennoi[i] = vecTOK[k];
			}
			
			if(Chast_to_Filtrate != 0)
			{
				ocmath_savitsky_golay(vec_Y_stepennoi, vec_Y_stepennoi_posrednik, vec_Y_stepennoi.GetSize(), vec_Y_stepennoi.GetSize()*Chast_to_Filtrate, -1, 5, 0, EDGEPAD_REFLECT);
				vec_Y_stepennoi = vec_Y_stepennoi_posrednik;
				for(i=0; i<vec_Y_stepennoi.GetSize(); i++, indFilt++) vecFiltrate[indFilt] = vec_Y_stepennoi[i];
			}
			
			NLFitSession FitSession;
			FitSession.SetFunction("Gauss");
			FitSession.SetMaxNumIter(Num_iter);
			FitSession.SetData(vec_Y_stepennoi, vec_X_stepennoi);
			
			vector vParamValues(4), vErrors;
			FitSession.ParamsInitValues();
			/*
			if(i_1 == 0) FitSession.ParamsInitValues();
			else 
			{
				vParamValues[0] = vecA[vec_Otr_Start_Indxs.GetSize()-1];
				vParamValues[1] = vecB[vec_Otr_Start_Indxs.GetSize()-1];
				vParamValues[2] = vecC[vec_Otr_Start_Indxs.GetSize()-1];
				vParamValues[3] = vecD[vec_Otr_Start_Indxs.GetSize()-1];
				FitSession.SetParamValues( vParamValues );
			}
			*/
			
			int nOutcome;
			FitSession.Fit(&nOutcome, false, false);
			
			FitSession.GetFitResultsParams(vParamValues, vErrors);
			
			//FitSession.GetYFromX(vec_X_stepennoi, vec_Y_stepennoi, vec_Y_stepennoi.GetSize());
			
			up_down = (vParamValues[3] >= 0) ? 0 : 1;
			if((vParamValues[2] < 0 || vParamValues[2] > 1000) && i_1!=0) vParamValues[2] = vecB[i_1-1];
			
			for (i = 0; i < vec_sokrashPILA.GetSize(); i++, lolo++)
			{
				vecOBRAB_TOKA[lolo] = vec_Y_stepennoi[i] = fxGAUSS(vec_X_stepennoi[i],vParamValues[0],vParamValues[1],vParamValues[2],vParamValues[3]);
			}
			
			vector v_whm_w_peak(10);
			/*
			вектор v_whm_w_peak:
				whm
				-[0] ширина на полувысоте
				-[1] Y на полувысоте
				-[2] X на полувысоте слева
				-[3] X на полувысоте справа
				W
				-[4] значение коэффициента W
				-[5] Y при W
				-[6] X при W слева
				-[7] X при W справа
				Peak
				-[8] Y пика
				-[9] X пика
			*/
			v_whm_w_peak = find_whm_Wparam_LRmidindxs(vec_X_stepennoi, vec_Y_stepennoi, 1, vParamValues[2], up_down);
			
			vecA[i_1] = v_whm_w_peak[8]; //Max Value
			vecB[i_1] = v_whm_w_peak[4]; //Temp
			vecC[i_1] = v_whm_w_peak[9]; //Peak Voltage
			if(v_whm_w_peak[0] < 1000 || i_1 == 0) 
				vecD[i_1] = v_whm_w_peak[0]; //Energy
			else if(i_1 != 0)
				vecD[i_1] = vecD[i_1-1]; //Energy
				
			break;
	}
}

vector find_whm_Wparam_LRmidindxs(vector &vX, vector &vY, int Diff_Gauss, float W, int up_down)
{
	vector v_whm_w_peak(10);
	/*
	вектор v_whm_w_peak:
		whm
		-[0] ширина на полувысоте
		-[1] Y на полувысоте
		-[2] X на полувысоте слева
		-[3] X на полувысоте справа
		W
		-[4] значение коэффициента W
		-[5] Y при W
		-[6] X при W слева
		-[7] X при W справа
		Peak
		-[8] Y пика
		-[9] X пика
	*/
	
	bool was_negative = false;
	if(Diff_Gauss == 1) was_negative = turn_tok_upwards(vY, 1, up_down);

	Curve crv(vX,vY);
	
	/* Peak */
	v_whm_w_peak[8] = max(vY);
	v_whm_w_peak[9] = xatymax(crv);

	if(Diff_Gauss == 0)
	{
		int i = 0, j = vY.GetSize()-1;
		float Y_whm = (max(vY) - min(vY))/2 + min(vY);
		
		while(vY[i] < Y_whm)
		{
			if(i >= vY.GetSize()-1) break;
			i++;
		}
		while(vY[j] < Y_whm)
		{
			if(j <= 0) break;
			j--;
		}
		/* whm */
		v_whm_w_peak[0] = abs(vX[i] - vX[j]);
		v_whm_w_peak[1] = Y_whm;
		v_whm_w_peak[2] = vX[i];
		v_whm_w_peak[3] = vX[j];
		/* W */
		v_whm_w_peak[4] = 0;
		v_whm_w_peak[5] = 0;
		v_whm_w_peak[6] = 0;
		v_whm_w_peak[7] = 0;
	}
	else
	{
		/* whm */
		v_whm_w_peak[0] = fwhm(crv, min(vY));
		v_whm_w_peak[1] = Curve_yfromX(&crv, v_whm_w_peak[9] - v_whm_w_peak[0]/2);
		v_whm_w_peak[2] = v_whm_w_peak[9] - v_whm_w_peak[0]/2;
		v_whm_w_peak[3] = v_whm_w_peak[9] + v_whm_w_peak[0]/2;
		/* W */
		v_whm_w_peak[4] = W;
		v_whm_w_peak[5] = Curve_yfromX(&crv, v_whm_w_peak[9] - W/2);
		v_whm_w_peak[6] = v_whm_w_peak[9] - W/2;
		v_whm_w_peak[7] = v_whm_w_peak[9] + W/2;
	}
	
	if(was_negative) 
	{
		vY *= -1;
		v_whm_w_peak[1] *= -1;
		v_whm_w_peak[5] *= -1;
		v_whm_w_peak[8] *= -1;
	}
	
	return v_whm_w_peak;
}

int pereschet_otrezka_zond(int &num_col, Worksheet &wks_graph_parent, Worksheet &wks_Legend)
{
	double Chast_to_Filtrate = 0;
	int index_nach_otrezka, index_konca_otrezka, size_otrezka, resistance, A_fix_or_not, B_fix_or_not, C_fix_or_not, D_fix_or_not, Filtered_or_Not=0, diagnostic_num, He_or_Ar;
	Worksheet wks_OrigData, wks_sheet1, wks_sheet2, wks_sheet3, wks_sheet4;
	if( Get_OriginData_wks(wks_graph_parent, wks_OrigData, wks_Legend) == 0 ) return 0;
	if(wks_Legend.Columns(0).GetUpperBound() < 3)
	{
		MessageBox(GetWindow(), "Невозможно пересчитать, на графике больше одного отрезка");
		return 0;
	}
	
	string str;
    wks_Legend.GetCell(6, 1, str);
    if(str.Match("He")) He_or_Ar=0;
    else He_or_Ar=1;
    
	is_str_numeric_integer(get_str_from_str_list(wks_graph_parent.Columns(0).GetUnits(), " ", 2), &resistance);
	is_str_numeric_integer(get_str_from_str_list(wks_Legend.Columns(1).GetLongName(), " ", 1), &index_nach_otrezka);
	is_str_numeric_integer(get_str_from_str_list(wks_Legend.Columns(1).GetLongName(), " ", 3), &index_konca_otrezka);
	size_otrezka = index_konca_otrezka - index_nach_otrezka + 1;
	
	Page pg =  wks_graph_parent.GetPage();
    wks_sheet1 = WksIdxByName(pg, "OrigData");
    wks_sheet2 = WksIdxByName(pg, "Fit");
    wks_sheet3 = WksIdxByName(pg, "Parameters");
    wks_sheet4 = WksIdxByName(pg, "Filtration");
    num_col = ColIdxByUnit(wks_sheet1, (string)index_nach_otrezka).GetIndex();
	Dataset dsPila(wks_sheet1.Columns(0));
	
	is_str_numeric(get_str_from_str_list(wks_sheet2.Columns(num_col).GetUnits(), " ", 1), &otsech_nachala);
    is_str_numeric(get_str_from_str_list(wks_sheet2.Columns(num_col).GetUnits(), " ", 3), &otsech_konca);
	
	if(pg.GetLongName().Match("Зонд1*")) diagnostic_num=0;
	if(pg.GetLongName().Match("Зонд2*")) diagnostic_num=1;
	
	GETN_BOX( treeTest );
    GETN_BEGIN_BRANCH(Start, "Fitting Options") 
    
    string strPar;
    wks_sheet3.GetCell(num_col-1, 1, strPar);
    is_str_numeric(strPar, &par_A);
    wks_sheet3.GetCell(num_col-1, 2, strPar);
    is_str_numeric(strPar, &par_B);
    wks_sheet3.GetCell(num_col-1, 3, strPar);
    is_str_numeric(strPar, &par_C);
    wks_sheet3.GetCell(num_col-1, 4, strPar);
    is_str_numeric(strPar, &par_D);
    
    GETN_MULTI_COLS_BRANCH(2, DYNA_MULTI_COLS_COMAPCT) GETN_OPTION_BRANCH(GETNBRANCH_OPEN)
		
		GETN_NUM(decayT3, "Коэффициент A", par_A)//GETN_EDITOR_SIZE_USER_OPTION("#5")
		GETN_CHECK(Check1, "Fix A", 0)
		GETN_NUM(decayT4, "Коэффициент B", par_B)//GETN_EDITOR_SIZE_USER_OPTION("#5")
		GETN_CHECK(Check2, "Fix B", 0)
		GETN_NUM(decayT5, "Коэффициент C", par_C)//GETN_EDITOR_SIZE_USER_OPTION("#5")
		GETN_CHECK(Check3, "Fix C", 0)
		GETN_NUM(decayT6, "Коэффициент D", par_D)//GETN_EDITOR_SIZE_USER_OPTION("#5")
		GETN_CHECK(Check4, "Fix D", 0)
		
    GETN_END_BRANCH(Start)
	
    GETN_SLIDEREDIT(decayT11, "Часть данных от начала которые будут линейно аппроксимированны", chast_to_fitln, "0.00|0.99|99")
    
    //GETN_SLIDEREDIT(decayT14, "Часть точек, которые будут отсечены от начала", otsech_nachala, "0.01|0.49|48")
    GETN_SLIDEREDIT(decayT15, "Часть точек, которые будут отсечены от конца", otsech_konca, "0.01|0.49|48")

	if(wks_sheet4) is_str_numeric(wks_sheet4.Columns(num_col).GetUnits(), &Chast_to_Filtrate);
    if(!Chast_to_Filtrate || Chast_to_Filtrate == NANUM) Chast_to_Filtrate = 0;
	if(Chast_to_Filtrate !=0) Filtered_or_Not = 1;
    else Chast_to_Filtrate = 0.01;

    GETN_BEGIN_BRANCH(Details, "Фильтрация данных") GETN_CHECKBOX_BRANCH(Filtered_or_Not)
    if(Filtered_or_Not == 1) GETN_OPTION_BRANCH(GETNBRANCH_OPEN)
    GETN_SLIDEREDIT(decayT13, "Часть точек фильтрации", Chast_to_Filtrate, "0.01|0.99|98") 
    GETN_END_BRANCH(Details)
    
	if(0==GetNBox( treeTest, "Пересчет этого отрезка" )) return 0;
	
	TreeNode tn1;
    treeTest.GetNode("Start").GetNode("decayT3").GetValue(par_A);
    treeTest.GetNode("Start").GetNode("decayT4").GetValue(par_B);
	treeTest.GetNode("Start").GetNode("decayT5").GetValue(par_C);
	treeTest.GetNode("Start").GetNode("decayT6").GetValue(par_D);
	treeTest.GetNode("decayT11").GetValue(chast_to_fitln);
	
	//treeTest.GetNode("decayT14").GetValue(otsech_nachala);
	treeTest.GetNode("decayT15").GetValue(otsech_konca);
	//otsech_nachala = (otsech_nachala == 0) ? 0.01 : otsech_nachala;
	otsech_konca = (otsech_konca == 0) ? 0.01 : otsech_konca;
	
	treeTest.GetNode("Start").GetNode("Check1").GetValue(A_fix_or_not);
	treeTest.GetNode("Start").GetNode("Check2").GetValue(B_fix_or_not);
	treeTest.GetNode("Start").GetNode("Check3").GetValue(C_fix_or_not);
	treeTest.GetNode("Start").GetNode("Check4").GetValue(D_fix_or_not);
	if(treeTest.Details.Use)
	{
		treeTest.GetNode("Details").GetNode("decayT13").GetValue(Chast_to_Filtrate);
		Chast_to_Filtrate = Chast_to_Filtrate/2;
	}
	else Chast_to_Filtrate = 0;
	
    vector vParams(4), vec_X_stepennoi(wks_sheet1.Columns(0)), vec_Y_stepennoi(wks_sheet1.Columns(num_col)), vec_Y_stepennoi_posrednik;
    int new_size = size_otrezka*(1 - otsech_nachala - otsech_konca), ind_X = vec_X_stepennoi.GetSize(), ind_Y = vec_Y_stepennoi.GetSize();
    
    if(new_size < ind_Y)
    {
		vec_Y_stepennoi.SetSize(new_size);
		vec_X_stepennoi.SetSize(new_size);
		wks_sheet1.Columns(num_col).SetNumRows(new_size);
		wks_sheet2.Columns(num_col).SetNumRows(new_size);
    }
    else if(new_size > ind_Y)
    {
    	float shag = vec_X_stepennoi[ind_X-1] - vec_X_stepennoi[ind_X-2];
    	vec_Y_stepennoi.SetSize(new_size);
    	vec_X_stepennoi.SetSize(new_size);
    	for(int i = ind_X; i < vec_X_stepennoi.GetSize(); i++)
    		vec_X_stepennoi[i] = vec_X_stepennoi[i-1] + shag;
    	
    	Column colLol;
    	if(diagnostic_num == 0) colLol = ColIdxByName(wks_OrigData, s_tz);
    	if(diagnostic_num == 1) colLol = ColIdxByName(wks_OrigData, s_tz2);
    	if(!colLol) 
    	{
    		MessageBox(GetWindow(), FuncErrors(222));
    		return 0;
    	}
    	vector vec_Tok_Orig(colLol);
    	
    	for(int j = 0, k = index_nach_otrezka + size_otrezka*otsech_nachala; j < vec_Y_stepennoi.GetSize(); j++, k++)
    		vec_Y_stepennoi[j] = vec_Tok_Orig[k]/-resistance;
    }
    
    vec_Y_stepennoi_posrednik.SetSize(vec_Y_stepennoi.GetSize());
    vec_Y_stepennoi_posrednik = vec_Y_stepennoi;
    if(Chast_to_Filtrate != 0) 
    	ocmath_savitsky_golay(vec_Y_stepennoi, vec_Y_stepennoi_posrednik, vec_Y_stepennoi.GetSize(), vec_Y_stepennoi.GetSize()*Chast_to_Filtrate, -1, 5, 0, EDGEPAD_REFLECT);
	//vec_Y_stepennoi = vec_Y_stepennoi_posrednik;
	
    NLFitSession FitSession;
    FitSession.SetFunction("Stepennaya_BAX");
	FitSession.SetMaxNumIter(400);
	
    vParams[0] = par_A;
	vParams[1] = par_B;
	vParams[2] = (par_C >= 10^5 || par_C <= -10^5) ? 1 : par_C;
	vParams[3] = (par_D >= 1000 || par_D <= 0) ? 10 : par_D;
			
	FitSession.SetData(vec_Y_stepennoi_posrednik, vec_X_stepennoi);
			
	if(A_fix_or_not == 0 && B_fix_or_not == 0 && chast_to_fitln != 0) koeff_fit_linear_stepennaya(vParams, vec_Y_stepennoi_posrednik, vec_X_stepennoi);
		
	FitSession.SetParamValues(vParams);
	
	if(chast_to_fitln != 0 || A_fix_or_not == 1) FitSession.SetParamFix(0, true);
	if(chast_to_fitln != 0 || B_fix_or_not == 1) FitSession.SetParamFix(1, true);
	if(C_fix_or_not == 1) FitSession.SetParamFix(2, true);
	if(D_fix_or_not == 1) FitSession.SetParamFix(3, true);
			
	int nOutcome;
	FitSession.Fit(&nOutcome, false, false);
			
	vector vParamValues, vErrors;
	FitSession.GetFitResultsParams(vParamValues, vErrors);
	
	Dataset ds_A(wks_sheet3.Columns(1)), ds_B(wks_sheet3.Columns(2)), ds_C(wks_sheet3.Columns(3)), ds_D(wks_sheet3.Columns(4)), dsObr_Tok(wks_sheet2.Columns(num_col)), 
			dsOrig_data(wks_sheet1.Columns(num_col)), dsFilted_data, dsDens(wks_sheet3.Columns(5));
	dsOrig_data = vec_Y_stepennoi;
	
	int nR1, nR2;
	if(wks_sheet4 && Chast_to_Filtrate != 0)
	{
		dsFilted_data.Attach(wks_sheet4.Columns(num_col));
		dsFilted_data = vec_Y_stepennoi_posrednik;
		//dsPila.Attach(wks_sheet4.Columns(0));
		//dsPila = vec_X_stepennoi;
	}
	else if(Chast_to_Filtrate != 0)
	{
		pg.AddLayer("Filtration");
		wks_sheet4 = pg.Layers(3);
		while( wks_sheet4.DeleteCol(0) );
		wks_sheet1.GetBounds(0, 1, nR1, -1, false);
		wks_sheet3.GetBounds(0, 1, nR2, -1, false);
		wks_sheet4.SetSize(nR1+1,nR2+2);
		wks_sheet4.Columns(0).SetLongName("pila");
		wks_sheet4.Columns(0).SetComments("0");
		wks_sheet4.Columns(0).SetType(OKDATAOBJ_DESIGNATION_X);
		wks_sheet4.Columns(0).SetUnits(wks_sheet1.Columns(0).GetUnits());
		Dataset ds_STOLBOV(wks_sheet4.Columns(0));
		ds_STOLBOV=vec_X_stepennoi;
		
		Tree trFormat;  
		trFormat.Root.CommonStyle.Fill.FillColor.nVal = 18; 
		DataRange dr;
		for(int i_2=1; i_2<nR2+2; i_2++)
		{
			//wks_sheet4.Columns(i_2).SetUnits(wks_sheet1.Columns(i_2).GetUnits());
			wks_sheet4.Columns(i_2).SetLongName(wks_sheet1.Columns(i_2).GetLongName());
			wks_sheet4.Columns(i_2).SetComments(wks_sheet1.Columns(i_2).GetComments());
			if(i_2 % 2 != 0)
			{
				dr.Add("X", wks_sheet4, 0, i_2, -1, i_2);
				if (0==dr.UpdateThemeIDs(trFormat.Root)) bool bRet = dr.ApplyFormat(trFormat, true, true);
				dr.Reset();
			}
		}
		dsFilted_data.Attach(wks_sheet4.Columns(num_col));
		dsFilted_data = vec_Y_stepennoi_posrednik;
		
		wks_sheet4.SetIndex(1);
	}
	else if(wks_sheet4)
	{
		dsFilted_data.Attach(wks_sheet4.Columns(num_col));
		dsFilted_data.SetSize(0);
	}
	
	if(Chast_to_Filtrate != 0)
	{
		string strFilt;
		strFilt.Format("%.2f", Chast_to_Filtrate*2);
		wks_sheet4.Columns(num_col).SetUnits(strFilt);
	}
	else if(wks_sheet4) wks_sheet4.Columns(num_col).SetUnits((string)0);
	
	string otsecheniya;
	otsecheniya.Format("S %.2f E %.2f", otsech_nachala, otsech_konca);
	wks_sheet2.Columns(num_col).SetUnits(otsecheniya);
	
	if(ind_X<vec_X_stepennoi.GetSize())
	{
		dsPila.Attach(wks_sheet1.Columns(0));
		dsPila = vec_X_stepennoi;
		dsPila.Attach(wks_sheet2.Columns(0));
		dsPila = vec_X_stepennoi;
	}
	
	ds_A[num_col-1]=vParamValues[0];
	ds_B[num_col-1]=vParamValues[1];
	ds_C[num_col-1]=vParamValues[2];
	ds_D[num_col-1]=vParamValues[3];
	
	if(diagnostic_num == 0) dens(false, num_col, dsDens, ds_A, ds_B, ds_C, ds_D, vec_X_stepennoi, He_or_Ar, S1);
	else dens(false, num_col, dsDens, ds_A, ds_B, ds_C, ds_D, vec_X_stepennoi, He_or_Ar, S2);
	
	if(FitSession.GetYFromX(vec_X_stepennoi, vec_Y_stepennoi_posrednik, vec_Y_stepennoi_posrednik.GetSize())) dsObr_Tok = vec_Y_stepennoi_posrednik;
		
	FitSession.SetParamFix(0, false);
	FitSession.SetParamFix(1, false);
	FitSession.SetParamFix(2, false);
	FitSession.SetParamFix(3, false);
	
    wks_sheet1.GetBounds(nR1, 1, nR2, -1, false);
    wks_sheet1.SetSize(nR2+1, -1);
    wks_sheet2.SetSize(nR2+1, -1);
    if(wks_sheet4) wks_sheet4.SetSize(nR2+1, -1);
	
	return -1;
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

Worksheet WksIdxByName(Page pg, string strLabel)
{
	foreach(Layer wks in pg.Layers)
	{
		if ( wks.GetName().CompareNoCase(strLabel) == 0 ) return wks;
	}
}

Column ColIdxByName(Worksheet wks, string strLabel)
{
	foreach(Column col in wks.Columns)
	{
		if ( col.GetLongName().CompareNoCase(strLabel) == 0 && ( col.GetUpperBound() > 0 ) ) return col;
	}
}

Column ColIdxByUnit(Worksheet wks, string strLabel)
{
	foreach(Column col in wks.Columns)
	{
		if (get_str_from_str_list(col.GetUnits(), " ", 0).CompareNoCase(strLabel) == 0) return col;
	}
}

void pererisovka_grafika_zond(GraphLayer graph, int num_col, Worksheet wks_graph_parent, Worksheet wks_Legend)
{
	Worksheet wks_sheet1, wks_sheet2, wks_sheet3, wks_sheetLegend, wks_sheet4;
	Page pg =  wks_graph_parent.GetPage(), pgLegend =  wks_Legend.GetPage();
	wks_sheet1 = WksIdxByName(pg, "OrigData");
	wks_sheet2 = WksIdxByName(pg, "Fit");
    wks_sheet3 = WksIdxByName(pg, "Parameters");
    wks_sheet4 = WksIdxByName(pg, "Filtration");
    wks_sheetLegend = pgLegend.Layers(1);
    
    Dataset dsTok_Orig(wks_sheet1.Columns(num_col)), dsX(wks_sheetLegend.Columns(0)), dsY(wks_sheetLegend.Columns(1)), dsSTOLBA(wks_sheet1.Columns(0)), dsA(wks_sheet3.Columns(1)), dsB(wks_sheet3.Columns(2)), dsC(wks_sheet3.Columns(3)), dsD(wks_sheet3.Columns(4)), dsDens(wks_sheet3.Columns(5));
    double PeresechX, He_or_Ar;
    int nPlots = graph.DataPlots.Count();
    
    if(wks_sheet4)
	{
		if(wks_sheet4.Columns(num_col).GetUpperBound() > 0)
		{
			if(nPlots <= 5)
			{
				DataRange dr5;
				dr5.Add(wks_sheet4, 0, "X");
				dr5.Add(wks_sheet4, num_col, "Y");
				graph.AddPlot(dr5, IDM_PLOT_LINE);
				DataPlot dp5 = graph.DataPlots(5);
				Tree tr4;
				tr4 = dp5.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
				tr4.Root.Line.Width.dVal=2.5;
				tr4.Root.Line.Color.nVal=2;
				dp5.ApplyFormat(tr4, true, true);
				vector<uint> vecUI = {0,5,1,2,3,4};
				graph.ReorderPlots(vecUI);
			}
		}
		else if(nPlots == 6) graph.RemovePlot(1);
	}

    string str;
    wks_Legend.GetCell(6, 1, str);
    if(str.Match("He"))He_or_Ar=0;
    else He_or_Ar=1;
    
	//PeresechX=dens(false, num_col, dsDens, dsA, dsB, dsC, dsD, dsSTOLBA, dsTok_Orig.GetSize(), He_or_Ar);
	dsTok_Orig.Attach(wks_sheet2.Columns(num_col));
    Curve crv(dsSTOLBA,dsTok_Orig);
	PeresechX = Curve_xfromY(&crv,0);
	if(PeresechX > dsSTOLBA[dsSTOLBA.GetSize()-1]) PeresechX = dsSTOLBA[dsSTOLBA.GetSize()-1];
	else if(PeresechX < dsSTOLBA[0]) PeresechX = dsSTOLBA[0];

    dsX[0] = dsSTOLBA[0];
    dsX[1] = dsSTOLBA[dsSTOLBA.GetSize()-1];
    dsY[0] = dsX[0]*dsB[num_col-1] + dsA[num_col-1];
    dsY[1] = dsX[1]*dsB[num_col-1] + dsA[num_col-1];

    dsX.Attach(wks_sheetLegend.Columns(2));
    dsY.Attach(wks_sheetLegend.Columns(3));

	dsSTOLBA.Attach(wks_sheet1.Columns(num_col));
				
    dsX[0] = PeresechX;
    dsX[1] = PeresechX;
    dsY[0] = min(dsSTOLBA);
    dsY[1] = max(dsSTOLBA);

	graph.Rescale();
	GraphLayer gl1 = graph.GetPage().Layers(0);
	legend_update(gl1, ALM_LONGNAME);
	
	dsSTOLBA.Attach(wks_sheet1.Columns(0));
	Axis axisY = gl1.YAxis, axisX = gl1.XAxis;
	if(dsSTOLBA[0] >= 0) axisX.Scale.From.dVal = 0;
	else axisX.Scale.From.dVal = dsSTOLBA[0];
	axisX.Scale.To.dVal = dsSTOLBA[dsSTOLBA.GetSize()-1];
	
	wks_Legend.SetCell(0, 1, dsA[num_col-1]);
	wks_Legend.SetCell(1, 1, dsB[num_col-1]);
	wks_Legend.SetCell(2, 1, dsC[num_col-1]);
	wks_Legend.SetCell(3, 1, dsD[num_col-1]);
	wks_Legend.SetCell(4, 1, dsDens[num_col-1]);
	wks_Legend.AutoSize();
	
	create_labels(0);
}

void graph()
{
	int zond_or_setoch;
	Page pg;
	string pgOrigName;
	Worksheet wks_test = Project.ActiveLayer();
    if(!wks_test)
    {
    	GraphLayer graph = Project.ActiveLayer();
    	if(!graph)
    	{
    		MessageBox(GetWindow(), "wks! || grapg!");
			return;
    	}
    	else 
    	{
    		pg = graph.GetPage();
			pgOrigName = pg.GetLongName();
			if(pgOrigName.Match("Зонд*")||pgOrigName.Match("Сетка*")||pgOrigName.Match("Цилиндр*")||pgOrigName.Match("Магнит*"))
			{
				if(pgOrigName.Match("Зонд*"))zond_or_setoch=0;
				if(pgOrigName.Match("Сетка*")||pgOrigName.Match("Цилиндр*")||pgOrigName.Match("Магнит*"))zond_or_setoch=1;
			}
			else
			{
				MessageBox(GetWindow(), "!Зонд или !Сетка или !Цилиндр или !Магнит (не тот лист)");
				return;
			}
    		
    		int num_col;
    		Worksheet wks_graph_parent, wks_Legend;
    		if(zond_or_setoch==0)
    		{
    			if(pereschet_otrezka_zond(num_col, wks_graph_parent, wks_Legend) == 0) return;
				pererisovka_grafika_zond(graph, num_col, wks_graph_parent, wks_Legend);
    		}
    		if(zond_or_setoch==1)
    		{
    			if(pereschet_otrezka_setoch(num_col, wks_graph_parent, wks_Legend) == 0) return;
    			pererisovka_grafika_setoch(graph, num_col, wks_graph_parent, wks_Legend);
    		}
    		return;
    	}
    }
    
    pg = wks_test.GetPage();
    pgOrigName = pg.GetLongName();
	if(pgOrigName.Match("Зонд*")||pgOrigName.Match("Сетка*")||pgOrigName.Match("Цилиндр*")||pgOrigName.Match("Магнит*"))
	{
		if(pgOrigName.Match("Зонд*"))zond_or_setoch=0;
		if(pgOrigName.Match("Сетка*")||pgOrigName.Match("Цилиндр*")||pgOrigName.Match("Магнит*"))zond_or_setoch=1;
	}
	else
	{
		MessageBox(GetWindow(), "!Зонд или !Сетка (не тот лист)");
		return;
	}
	
	if(zond_or_setoch==0)graphZond();
	if(zond_or_setoch==1)graphSetoch();
}

void graphZond()
{
	time_t start_time, finish_time;
	time(&start_time);
	
	Worksheet wks_test = Project.ActiveLayer();
    if(!wks_test)
    {
    	MessageBox(GetWindow(), "wks!");
    	return;
    }
    Page pg = wks_test.GetPage();
    string pgOrigName = pg.GetLongName(); 
    int SVCH_or_VCH=0, kolvo_col, first_c_i, last_c_i, r1, r2, comp_evo=-1;
    vector<int> vROWS, vCOLS, vr2, vc2;
    
    Worksheet wks_sheet1, wks_sheet2, wks_sheet3, wks_sheetLegend, wks_sheet4, wks_SevCols;
    wks_sheet1 = WksIdxByName(pg, "OrigData");
    wks_sheet2 = WksIdxByName(pg, "Fit");
    wks_sheet3 = WksIdxByName(pg, "Parameters");
    wks_sheet4 = WksIdxByName(pg, "Filtration");
    if(!wks_sheet1 || !wks_sheet2 || !wks_sheet3)
    {
    	MessageBox(GetWindow(), "Ошибка! Не найден лист с данными");
		return;
    }
    
    if(!wks_test.GetSelectedRange(vROWS, vCOLS, vr2, vc2))
    {
    	if(wks_test == wks_sheet3) MessageBox(GetWindow(), "!Selection (выберите строку)");
    	else MessageBox(GetWindow(), "!Selection (выберите колонку)");
		return;
    }
    if(!wks_test.GetSelection(first_c_i, last_c_i, r1, r2))
    {
    	if(wks_test == wks_sheet3) MessageBox(GetWindow(), "!Selection (выберите строку)");
    	else MessageBox(GetWindow(), "!Selection (выберите колонку)");
		return;
    }
	
    if(wks_test == wks_sheet3)
    {
    	if(vROWS.GetSize() >= r2-r1+1 || (vROWS.GetSize() <= r2-r1+1 && vROWS.GetSize() > 1)) 
    	{
    		for(int i = 0, j; i < vROWS.GetSize(); i++) 
    		{
    			for(j = i+1; j < vROWS.GetSize(); j++)
    			{
					if(vROWS[i] == vROWS[j])
					{
						vROWS.RemoveAt(j);
						j--;
					}
    			}
    		}
    		kolvo_col = vROWS.GetSize();
    	}
    	else
    	{
    		kolvo_col = r2-r1+1;
    		vROWS.SetSize(kolvo_col);
    		vROWS[0] = r1;
    		for(int i = 1; i < vROWS.GetSize(); i++) vROWS[i] = vROWS[i-1] + 1;
    	}
    	for(int i = 1; i < vROWS.GetSize(); i++) if(vROWS[i] == vROWS[i-1]) vROWS.RemoveAt(i);
    }
    else
    {
    	if(first_c_i == 0) first_c_i = 1;
    	if(vCOLS[0] == 0) vCOLS[0] = 1;
    	if(vCOLS.GetSize() >= last_c_i-first_c_i+1 || (vCOLS.GetSize() <= last_c_i-first_c_i+1 && vCOLS.GetSize() > 1))
    	{
    		for(int i = 0, j; i < vCOLS.GetSize(); i++)
    		{
    			for(j = i+1; j < vCOLS.GetSize(); j++)
    			{
					if(vCOLS[i] == vCOLS[j]) 
					{
						vCOLS.RemoveAt(j);
						j--;
					}
    			}
    		}
    		kolvo_col = vCOLS.GetSize();
    	}
    	else 
    	{ 
    		kolvo_col = last_c_i-first_c_i+1;
    		vCOLS.SetSize(kolvo_col);
    		vCOLS[0] = first_c_i;
    		for(int i = 1; i < vCOLS.GetSize(); i++) vCOLS[i] = vCOLS[i-1] + 1;
    	}
    }
    
    if(kolvo_col>1)
	{
		GETN_BOX( treeTest ); 
		GETN_RADIO_INDEX_EX(decayT1, "", 0, "Сравнить выбранные отрезки|Построить эволюцию выбранных отрезков") 
		if(0==GetNBox( treeTest, "График нескольких отрезков" )) return;
		treeTest.GetNode("decayT1").GetValue(comp_evo);
	}
    
    string str_SVCH_or_VCH = get_str_from_str_list(wks_sheet1.Columns(0).GetUnits(), " ", 0);
    string str_He_or_Ar = get_str_from_str_list(wks_sheet1.Columns(0).GetUnits(), " ", 1);
    
    Column colA=wks_sheet3.Columns(1), colB=wks_sheet3.Columns(2), colC=wks_sheet3.Columns(3), colD=wks_sheet3.Columns(4), colDens=wks_sheet3.Columns(5);
	Dataset dsA, dsB, dsC, dsD, dsDens, dsPILA(wks_sheet1.Columns(0)), dsSTOLBA, dsOD, dsFlt, dsDfn, dsSevT, dsTimeSynth(wks_sheet3.Columns(0));
	
	if(wks_test == wks_sheet3)
    {
		if(dsTimeSynth[vROWS[vROWS.GetSize()-1]] == 0)
		{
			MessageBox(GetWindow(), "Ошибка во времени! Последний отрезок обработан с ошибкой (время начала равно 0). "
				+ "Для решения проблемы удалите отрезки со временем начала 0, или пересчитайте импульс");
			return;
		}
	}
	else
	{
		if(dsTimeSynth[vCOLS[vCOLS.GetSize()-1]-1] == 0)
		{
			MessageBox(GetWindow(), "Ошибка во времени! Последний отрезок обработан с ошибкой (время начала равно 0). "
				+ "Для решения проблемы удалите отрезки со временем начала 0, или пересчитайте импульс");
			return;
		}
	}
	
    dsA.Attach(colA);
	dsB.Attach(colB);
	dsC.Attach(colC);
	dsD.Attach(colD);
	dsDens.Attach(colDens);
	vector vecA, vecB, vecMinY, vecMaxY, vecPeresechX;
	
	DataRange dr1,dr2,dr3,dr4,dr5;
	double time_step;
	string str, str1;
	int k, size_pily=dsPILA.GetSize(), num_col, k_end;
	if(comp_evo != 0)
	{
		if(wks_test == wks_sheet3)
		{
			if(kolvo_col>1)
			{
				int i;
				dsOD.Create(0);
				if(wks_sheet4) dsFlt.Create(0);
				dsDfn.Create(0);
				for(i = 0; i < vROWS.GetSize(); i++)
				{
					Dataset dsSlave;
					dsSlave.Attach(wks_sheet1.Columns(vROWS[i]+1));
					dsOD.Append(dsSlave, REDRAW_NONE);
					if(wks_sheet4)
					{
						dsSlave.Attach(wks_sheet4.Columns(vROWS[i]+1));
						if(dsSlave.GetUpperBound()<=0)
							dsFlt.SetSize(dsFlt.GetSize()+wks_sheet4.Columns(0).GetUpperBound());
						else
							dsFlt.Append(dsSlave, REDRAW_NONE);
					}
					dsSlave.Attach(wks_sheet2.Columns(vROWS[i]+1));
					dsDfn.Append(dsSlave, REDRAW_NONE);
				}
				
				dsSevT.Create(dsOD.GetSize());
				dsSevT[0] = dsTimeSynth[vROWS[0]];
				for(i = 1; i < dsSevT.GetSize(); i++) dsSevT[i] = dsSevT[i-1] + (dsTimeSynth[vROWS[vROWS.GetSize()-1]] + 
					(dsTimeSynth[vROWS[vROWS.GetSize()-1]]-dsTimeSynth[vROWS[vROWS.GetSize()-2]]) - dsTimeSynth[vROWS[0]])/dsSevT.GetSize();
			}
			else 
			{
				dr1.Add(wks_sheet1, 0, "X");
				dr2.Add(wks_sheet2, 0, "X");
				if(wks_sheet4) dr5.Add(wks_sheet4, 0, "X");
				num_col = vROWS[0]+1;
				if(wks_sheet4) dr5.Add(wks_sheet4, vROWS[0]+1, "Y");
				dr1.Add(wks_sheet1, vROWS[0]+1, "Y");
				vecA.SetSize(1);
				vecB.SetSize(1);
				vecA[0] = dsA[vROWS[0]];
				vecB[0] = dsB[vROWS[0]];
				dsSTOLBA.Attach(wks_sheet1.Columns(vROWS[0]+1));
				dr2.Add(wks_sheet2, vROWS[0]+1, "Y");
				vecMinY.SetSize(1);
				vecMaxY.SetSize(1);
				vecPeresechX.SetSize(1);
				vecMinY[0] = min(dsSTOLBA);
				vecMaxY[0] = max(dsSTOLBA);
				//vecPeresechX[0]=dens(false, vROWS[0]+1, dsDens, dsA, dsB, dsC, dsD, dsPILA, dsPILA.GetSize(), 0, S);
				dsSTOLBA.Attach(wks_sheet2.Columns(vROWS[0]+1));
				Curve crv(dsPILA,dsSTOLBA);
				vecPeresechX[0] = Curve_xfromY(&crv,0);
				if(vecPeresechX[0] > dsPILA[dsPILA.GetSize()-1]) vecPeresechX[0] = dsPILA[dsPILA.GetSize()-1];
				else if(vecPeresechX[0] < dsPILA[0]) vecPeresechX[0] = dsPILA[0];
			}
		}
		else
		{
			if(kolvo_col>1)
			{
				int i;
				dsOD.Create(0);
				if(wks_sheet4) dsFlt.Create(0);
				dsDfn.Create(0);
				for(i = 0; i < vCOLS.GetSize(); i++)
				{
					Dataset dsSlave;
					dsSlave.Attach(wks_sheet1.Columns(vCOLS[i]));
					dsOD.Append(dsSlave, REDRAW_NONE);
					if(wks_sheet4)
					{
						dsSlave.Attach(wks_sheet4.Columns(vCOLS[i]));
						if(dsSlave.GetUpperBound()<=0)
							dsFlt.SetSize(dsFlt.GetSize()+wks_sheet4.Columns(0).GetUpperBound());
						else
							dsFlt.Append(dsSlave, REDRAW_NONE);
					}
					dsSlave.Attach(wks_sheet2.Columns(vCOLS[i]));
					dsDfn.Append(dsSlave, REDRAW_NONE);
				}
				
				dsSevT.Create(dsOD.GetSize());
				dsSevT[0] = dsTimeSynth[vCOLS[0]-1];
				for(i = 1; i < dsSevT.GetSize(); i++) dsSevT[i] = dsSevT[i-1] + (dsTimeSynth[vCOLS[vCOLS.GetSize()-1]-1] + 
					(dsTimeSynth[vCOLS[vCOLS.GetSize()-1]-1]-dsTimeSynth[vCOLS[vCOLS.GetSize()-1]-2]) - dsTimeSynth[vCOLS[0]-1])/dsSevT.GetSize();
			}
			else 
			{
				dr1.Add(wks_sheet1, 0, "X");
				dr2.Add(wks_sheet2, 0, "X");
				if(wks_sheet4) dr5.Add(wks_sheet4, 0, "X");
				num_col = vCOLS[0];
				if(wks_sheet4) dr5.Add(wks_sheet4, vCOLS[0], "Y");
				dr1.Add(wks_sheet1, vCOLS[0], "Y");
				vecA.SetSize(1);
				vecB.SetSize(1);
				vecA[0] = dsA[vCOLS[0]-1];
				vecB[0] = dsB[vCOLS[0]-1];
				dsSTOLBA.Attach(wks_sheet1.Columns(vCOLS[0]));
				dr2.Add(wks_sheet2, vCOLS[0], "Y");
				vecMinY.SetSize(1);
				vecMaxY.SetSize(1);
				vecPeresechX.SetSize(1);
				vecMinY[0] = min(dsSTOLBA);
				vecMaxY[0] = max(dsSTOLBA);
				//vecPeresechX[0]=dens(false, vCOLS[0], dsDens, dsA, dsB, dsC, dsD, dsPILA, dsPILA.GetSize(), 0);
				dsSTOLBA.Attach(wks_sheet2.Columns(vCOLS[0]));
				Curve crv(dsPILA,dsSTOLBA);
				vecPeresechX[0] = Curve_xfromY(&crv,0);
				if(vecPeresechX[0] > dsPILA[dsPILA.GetSize()-1]) vecPeresechX[0] = dsPILA[dsPILA.GetSize()-1];
				else if(vecPeresechX[0] < dsPILA[0]) vecPeresechX[0] = dsPILA[0];
			}
		}
	}
	
	if(wks_test == wks_sheet3)
	{
		if(kolvo_col>1)
		{
			str = get_str_from_str_list(wks_sheet1.Columns(min(vROWS)+1).GetLongName(), " ", 0) + " - " + get_str_from_str_list(wks_sheet1.Columns(max(vROWS)+1).GetLongName(), " ", 2);
			str1 = get_str_from_str_list(wks_sheet1.Columns(min(vROWS)+1).GetUnits(), " ", 0) + " - " + get_str_from_str_list(wks_sheet1.Columns(max(vROWS)+1).GetUnits(), " ", 2);
			k=vROWS[0];
			k_end=vROWS[vROWS.GetSize()-1];
		}
		else 
		{
			str = wks_sheet1.Columns(vROWS[0]+1).GetLongName();
			str1 = wks_sheet1.Columns(vROWS[0]+1).GetUnits();
			k=vROWS[0];
		}
	}
	else
	{
		if(kolvo_col>1)
		{
			str = get_str_from_str_list(wks_sheet1.Columns(min(vCOLS)).GetLongName(), " ", 0) + " - " + get_str_from_str_list(wks_sheet1.Columns(max(vCOLS)).GetLongName(), " ", 2);
			str1 = get_str_from_str_list(wks_sheet1.Columns(min(vCOLS)).GetUnits(), " ", 0) + " - " + get_str_from_str_list(wks_sheet1.Columns(max(vCOLS)).GetUnits(), " ", 2);
			k=vCOLS[0]-1;
			k_end=vCOLS[vCOLS.GetSize()-1]-1;
		}
		else 
		{
			str = wks_sheet1.Columns(vCOLS[0]).GetLongName();
			str1 = wks_sheet1.Columns(vCOLS[0]).GetUnits();
			k=vCOLS[0]-1;
		}
	}
    
	GraphPage gp;
	gp.Create("Origin");
	GraphLayer gl1 = gp.Layers(0);
	wks_sheet1.ConnectTo(gl1);
	Tree tr;
	tr = gl1.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
	tr.Root.Legend.Font.Size.nVal = 12;
	if(0 == gl1.UpdateThemeIDs(tr.Root) ) bool bRet = gl1.ApplyFormat(tr, true, true);
	
	Axis axesX = gl1.XAxis, axesY = gl1.YAxis ;	
	axesX.Additional.ZeroLine.nVal = 1;
	axesY.Additional.ZeroLine.nVal = 1;
    
	if(kolvo_col>1) gp.SetLongName ( FindPageNameNumber ( get_str_from_str_list(pgOrigName, " ", 0) + " " + get_str_from_str_list(pgOrigName, " ", 3) + '(' + get_str_from_str_list(pgOrigName, " ", 4) + ')' + " Отрезки " + (k+1) + " - " + (k_end+1) ) );
	else gp.SetLongName ( FindPageNameNumber ( get_str_from_str_list(pgOrigName, " ", 0) + " " + get_str_from_str_list(pgOrigName, " ", 3) + '(' + get_str_from_str_list(pgOrigName, " ", 4) + ')' + " Отрезок " + (k+1) ) );
    if(comp_evo == 1) create_legend_zond(dsA[k],dsB[k],dsC[k],dsD[k],dsDens[k],str,str_SVCH_or_VCH,str_He_or_Ar,pgOrigName,str1,wks_sheetLegend, kolvo_col, wks_SevCols);
	else create_legend_zond(dsA[k],dsB[k],dsC[k],dsD[k],dsDens[k],str,str_SVCH_or_VCH,str_He_or_Ar,pgOrigName,str1,wks_sheetLegend, kolvo_col, NULL);
	if(comp_evo == 1 && kolvo_col > 1)
    {
    	Dataset dsSlave;
    	dsSlave.Create(dsOD.GetSize());
    	dsSlave.Attach(wks_SevCols.Columns(0));
    	dsSlave = dsOD;
    	if(wks_sheet4)
    	{
			dsSlave.Attach(wks_SevCols.Columns(1));
			dsSlave = dsFlt;
    	}
    	dsSlave.Attach(wks_SevCols.Columns(2));
    	dsSlave = dsDfn;
    	dsSlave.Attach(wks_SevCols.Columns(3));
    	dsSlave = dsSevT;
    }
    
    DataPlot dp1,dp2,dp3,dp4,dp5;
    Tree tr1, tr2, tr3, tr4;
    
    if(kolvo_col == 1)
    {
		Dataset dsX, dsY;
		dsX.Attach(wks_sheetLegend.Columns(0));
		dsY.Attach(wks_sheetLegend.Columns(1));
		dsX.SetSize(vecA.GetSize()*2);
		dsY.SetSize(vecA.GetSize()*2);
		for(int i=0, j=0; i<vecA.GetSize(); i++, j=j+2)
		{
			dsX[j] = dsPILA[0];
			dsX[j+1] = dsPILA[dsPILA.GetSize()-1];
			dsY[j] = dsX[j]*vecB[i] + vecA[i];
			dsY[j+1] = dsX[j+1]*vecB[i] + vecA[i];
		} 
		dsX.Attach(wks_sheetLegend.Columns(2));
		dsY.Attach(wks_sheetLegend.Columns(3));
		dsX.SetSize(vecMinY.GetSize()*2);
		dsY.SetSize(vecMaxY.GetSize()*2);
		for(int i_1=0, j_1=0; i_1<vecMinY.GetSize(); i_1++, j_1=j_1+2)
		{
			dsX[j_1] = vecPeresechX[i_1];
			dsX[j_1+1] = vecPeresechX[i_1];
			dsY[j_1] = vecMinY[i_1];
			dsY[j_1+1] = vecMaxY[i_1];
		}
		dr3.Add(wks_sheetLegend, 0, "X");
		dr3.Add(wks_sheetLegend, 1, "Y");
		dr4.Add(wks_sheetLegend, 2, "X");
		dr4.Add(wks_sheetLegend, 3, "Y");
		
		gl1.AddPlot(dr1, IDM_PLOT_LINE);
		gl1.AddPlot(dr2, IDM_PLOT_LINE);
		gl1.AddPlot(dr3, IDM_PLOT_LINE);
		gl1.AddPlot(dr4, IDM_PLOT_LINE);
		if(wks_sheet4 && wks_sheet4.Columns(num_col).GetUpperBound() > 0) gl1.AddPlot(dr5, IDM_PLOT_LINE);
		
		dp1 = gl1.DataPlots(0);
		dp2 = gl1.DataPlots(1); 
		dp3 = gl1.DataPlots(2);
		dp4 = gl1.DataPlots(3);
		if(wks_sheet4) dp5 = gl1.DataPlots(4);
		
		tr1 = dp1.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
		tr2 = dp2.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
		tr3 = dp3.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
		
		tr1.Root.Line.Width.dVal=2;
		tr1.Root.Line.Color.nVal = set_color_rnd();
		
		tr2.Root.Line.Width.dVal=2.5;
		tr2.Root.Line.Color.nVal=1;
		
		tr3.Root.Line.Width.dVal=1.1;
		tr3.Root.Line.Color.nVal=3;
		
		dp1.ApplyFormat(tr1, true, true);
		dp2.ApplyFormat(tr2, true, true);
		dp3.ApplyFormat(tr3, true, true);
		dp4.ApplyFormat(tr3, true, true);
		if(wks_sheet4 && wks_sheet4.Columns(num_col).GetUpperBound() > 0)
		{
			tr4 = dp5.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
			tr4.Root.Line.Width.dVal=2.5;
			tr4.Root.Line.Color.nVal=2;
			dp5.ApplyFormat(tr4, true, true);
			vector<uint> vecUI = {0,4,1,2,3};
			gl1.ReorderPlots(vecUI);
		}
    }
    
    if(comp_evo == 0 && kolvo_col > 1)
    {
    	Curve cc;
    	int i, k = 0, color = 0;
    	if(wks_test == wks_sheet3)
    	{
    		/////////////ORIGDATA/////////////
    		for(i = 0; i < vROWS.GetSize(); i++)
    		{
				cc.Attach(wks_sheet1, 0, vROWS[i]+1);
				gl1.AddPlot(cc, IDM_PLOT_SCATTER);
				tr1 = gl1.DataPlots(k).GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
				tr1.Root.Symbol.Size.nVal=2;
				gl1.DataPlots(k).ApplyFormat(tr1, true, true);
				color = set_color_except_bad(color);
				gl1.DataPlots(k).SetColor(color);//всего цветов 22
				k++;
    		}
    		/////////////SMOOTH/////////////
    		/*
			for(i = 0; i < vROWS.GetSize(); i++)
			{	
				if(wks_sheet4 && wks_sheet4.Columns(vROWS[i]+1).GetUpperBound() > 0)
				{
					cc.Attach(wks_sheet4, 0, vROWS[i]+1);
					gl1.AddPlot(cc, IDM_PLOT_LINE);
					tr1 = gl1.DataPlots(k).GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
					tr1.Root.Line.Width.dVal=2;
					gl1.DataPlots(k).ApplyFormat(tr1, true, true);
					gl1.DataPlots(k).SetColor(k+1);
					k++;
				}
			}
			*/
			/////////////FITTING/////////////
			for(i = 0; i < vROWS.GetSize(); i++)
    		{
				cc.Attach(wks_sheet2, 0, vROWS[i]+1);
				gl1.AddPlot(cc, IDM_PLOT_LINE);
				tr1 = gl1.DataPlots(k).GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
				tr1.Root.Line.Width.dVal=2;
				gl1.DataPlots(k).ApplyFormat(tr1, true, true);
				color = set_color_except_bad(color);
				gl1.DataPlots(k).SetColor(color);
				k++;
    		}
    	}
    	else
    	{
    		/////////////ORIGDATA/////////////
    		for(i = 0; i < vCOLS.GetSize(); i++)
    		{
				cc.Attach(wks_sheet1, 0, vCOLS[i]);
				gl1.AddPlot(cc, IDM_PLOT_SCATTER);
				tr1 = gl1.DataPlots(k).GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
				tr1.Root.Symbol.Size.nVal=2;
				gl1.DataPlots(k).ApplyFormat(tr1, true, true);
				color = set_color_except_bad(color);
				gl1.DataPlots(k).SetColor(color);
				k++;
    		}
    		/////////////SMOOTH/////////////
    		/*
			for(i = 0; i < vCOLS.GetSize(); i++)
			{
				if(wks_sheet4 && wks_sheet4.Columns(vCOLS[i]).GetUpperBound() > 0)
				{
					cc.Attach(wks_sheet4, 0, vCOLS[i]);
					gl1.AddPlot(cc, IDM_PLOT_LINE);
					tr1 = gl1.DataPlots(k).GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
					tr1.Root.Line.Width.dVal=2;
					gl1.DataPlots(k).ApplyFormat(tr1, true, true);
					gl1.DataPlots(k).SetColor(k+1);
					k++;
				}
			}
			*/
			/////////////FITTING/////////////
			for(i = 0; i < vCOLS.GetSize(); i++)
    		{
				cc.Attach(wks_sheet2, 0, vCOLS[i]);
				gl1.AddPlot(cc, IDM_PLOT_LINE);
				tr1 = gl1.DataPlots(k).GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
				tr1.Root.Line.Width.dVal=2;
				gl1.DataPlots(k).ApplyFormat(tr1, true, true);
				color = set_color_except_bad(color);
				gl1.DataPlots(k).SetColor(color);
				k++;
    		}
    	}
    }
    if(comp_evo == 1 && kolvo_col > 1)
    {
    	dr1.Add(wks_SevCols, 0, "Y");
    	dr1.Add(wks_SevCols, 3, "X");
		dr2.Add(wks_SevCols, 1, "Y");
		dr2.Add(wks_SevCols, 3, "X");
		dr3.Add(wks_SevCols, 2, "Y");
		dr3.Add(wks_SevCols, 3, "X");
		
		gl1.AddPlot(dr1, IDM_PLOT_LINE);
		if(wks_sheet4 && !dr2.IsEmpty()) gl1.AddPlot(dr2, IDM_PLOT_SCATTER);
		gl1.AddPlot(dr3, IDM_PLOT_SCATTER);
		
		dp1 = gl1.DataPlots(0);
		if(wks_sheet4 && !dr2.IsEmpty())
		{
			dp2 = gl1.DataPlots(1); 
			dp3 = gl1.DataPlots(2);
		}
		else dp3 = gl1.DataPlots(1);
		
		tr1 = dp1.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
		if(wks_sheet4 && !dr2.IsEmpty()) tr2 = dp2.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
		tr3 = dp3.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
		
		tr1.Root.Line.Width.dVal=1;
		tr1.Root.Line.Color.nVal=0;
		
		tr2.Root.Symbol.Size.nVal=2;
		tr2.Root.Symbol.EdgeColor.nVal=2;
		
		tr3.Root.Symbol.Size.nVal=2;
		tr3.Root.Symbol.EdgeColor.nVal=1;
		
		dp1.ApplyFormat(tr1, true, true);
		if(wks_sheet4 && !dr2.IsEmpty()) dp2.ApplyFormat(tr2, true, true);
		dp3.ApplyFormat(tr3, true, true);
    }
    
    Axis axisY = gl1.YAxis, axisX = gl1.XAxis;
	tr.Reset();
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
	
	if(comp_evo == 1) gl1.GraphObjects("XB").Text = "Время, с";
	else gl1.GraphObjects("XB").Text = "Напряжение, В";
	gl1.GraphObjects("YL").Text = "Ток, А";
	legend_update(gl1, ALM_LONGNAME);
	gl1.Rescale();
	
	if(comp_evo == 0 || kolvo_col == 1)
    {
		if(dsPILA[0] >= 0) axisX.Scale.From.dVal = 0;
		else axisX.Scale.From.dVal = dsPILA[0];
		axisX.Scale.To.dVal = dsPILA[dsPILA.GetSize()-1];
		axisX.Scale.IncrementBy.nVal = 0; // 0=increment by value; 1=number of major ticks
		axisX.Scale.Value.dVal = 20; // Increment value
		
		axisY.Scale.IncrementBy.nVal = 1; // 0=increment by value; 1=number of major ticks
		axisY.Scale.MajorTicksCount.dVal = 10;
		
		if(kolvo_col == 1) create_labels(0);
    }
    else 
    {
    	axisX.Scale.IncrementBy.nVal = 1; // 0=increment by value; 1=number of major ticks
		axisX.Scale.MajorTicksCount.dVal = 10;
		axisY.Scale.IncrementBy.nVal = 1; // 0=increment by value; 1=number of major ticks
		axisY.Scale.MajorTicksCount.dVal = 10;
    }
	
	time(&finish_time);
	printf("Run time graph %d secs\n", finish_time - start_time);
}

double dens(bool cycle_r_not, int num_col_to_do, vector &vecDens, vector &vecA, vector &vecB, vector &vecC, vector &vecD, 
			vector &vec_sokrashPILA, int He_or_Ar, double S)
{
	vecDens.SetSize(vecA.GetSize());
	
	double A,B,C,D,ans,Density;
	
	double x, x1=vec_sokrashPILA[0], x2=vec_sokrashPILA[vec_sokrashPILA.GetSize()-1];
	if(cycle_r_not)
	{
		for(int i=0; i < vecA.GetSize(); i++)
		{
			A=vecA[i];
			B=vecB[i];
			C=vecC[i];
			D=vecD[i];
			x = metod_hord(x1,x2,A,B,C,D);
			if(x==NANUM||x<=x1||x>=x2)
			{
				x = metod_Newton(x2,A,B,C,D);
				if(x==NANUM||x<=x1||x>=x2)
				{
					x=x1;
					int n=0;
					while(fx(x,A,B,C,D)<0)
					{
						x = x + 0.15;
						n++;
						if(x<=x1||x>=x2||n>1000) break;
					}
				}
			}
			ans=abs(A+B*x);
			if(He_or_Ar==0)Density=ans*(10^-6)/(0.52026*S*(1.602*10^-19)*sqrt((1.380649*10^-23)*D*11604.51812/M_He)); //моя формула
			//if(He_or_Ar==0)Density=ans*(10^-6)/(S*(1.602*10^-19)*sqrt(2*D*(1.602*10^-19)/M_He)); //формула сергея
			else Density=ans*(10^-6)/(0.52026*S*(1.602*10^-19)*sqrt((1.380649*10^-23)*D*11604.51812/M_Ar)); //моя формула
			vecDens[i]=Density;
		}
	}
	else
	{
		A=vecA[num_col_to_do-1];
		B=vecB[num_col_to_do-1];
		C=vecC[num_col_to_do-1];
		D=vecD[num_col_to_do-1];
		x = metod_hord(x1,x2,A,B,C,D);
		if(x==NANUM||x<=x1||x>=x2)
		{
			x = metod_Newton(x2,A,B,C,D);
			if(x==NANUM||x<=x1||x>=x2)
			{
				x=x1;
				int n=0;
				while(fx(x,A,B,C,D)<0)
				{
					x = x + 0.15;
					n++;
					if(x<=x1||x>=x2||n>1000) break;
				}
			}
		}
		ans=abs(A+B*x);
		if(He_or_Ar==0)Density=ans*(10^-6)/(0.52026*S*(1.602*10^-19)*sqrt((1.380649*10^-23)*D*11604.51812/M_He)); //моя формула
		//if(He_or_Ar==0)Density=ans*(10^-6)/(S*(1.602*10^-19)*sqrt(2*D*(1.602*10^-19)/M_He)); //формула сергея
		else Density=ans*(10^-6)/(0.52026*S*(1.602*10^-19)*sqrt((1.380649*10^-23)*D*11604.51812/M_Ar)); //моя формула
		vecDens[num_col_to_do-1]=Density;
	}
	
	return x;
}

double fxGAUSS(double x,double y0,double xc,double w,double A)
{
	return y0 + (A/(w*sqrt(pi/2)))*exp(-2*((x-xc)/w)^2);
}

double fx(double x,double A,double B,double C,double D)
{
	return A+B*x+C*exp(x/D);
}

double dfx(double x,double B,double C,double D)
{
	return B+(C*exp(x/D))/D;
}

double dfxdC(double x,double D)
{
	return exp(x/D);
}

double dfxdD(double x,double C,double D)
{
	return -(C*x*exp(x/D))/D^2;
}

double metod_hord(double x0, double x1, double A,double B,double C,double D)
{
	double x_cashe=2*x1;
	int n=0;
	while(abs(x1 - x0) > 10^-10)
	{
		x0 = x1 - (x1 - x0)*fx(x1,A,B,C,D)/(fx(x1,A,B,C,D) - fx(x0,A,B,C,D));
		x1 = x0 - (x0 - x1)*fx(x0,A,B,C,D)/(fx(x0,A,B,C,D) - fx(x1,A,B,C,D));
		n++;
		if(x1>x_cashe||x0>x_cashe||n>500) 
		{
			x1=x_cashe;
			break;
		}
	}
	return x1;
}

double metod_Newton(double x0, double A,double B,double C,double D)
{
	double x1, x_cashe=2*x0;
	int n=0;
	while(abs(x1 - x0) > 10^-10)
	{
		x1 = x0 - fx(x0,A,B,C,D)/dfx(x0,B,C,D);
		x0 = x1 - fx(x1,A,B,C,D)/dfx(x1,B,C,D);
		n++;
		if(x1>x_cashe||x0>x_cashe||n>500) 
		{
			x1=x_cashe;
			break;
		}
	}
	return x1;
}

void convert_time_to_pts(int v_tok_size, int chastota_discret_zonda)
{
	float tot_time = (float)v_tok_size/chastota_discret_zonda;
	
	float st_ot_tot = st_time_end_time[0]/tot_time;
	float end_ot_tot = st_time_end_time[1]/tot_time;
	
	if(st_time_end_time[0] <= tot_time && st_time_end_time[1] <= tot_time)
	{
		i_nach = v_tok_size * st_ot_tot;
		i_konec = v_tok_size * end_ot_tot;
	}
}

int find_otrezki_toka_i_obrab_pily(vector &vec_Otr_Start_Indxs, vector &vecTOK_ZONDA, int chastota_discret_zonda, int chastota_pily, vector &vecPILA_ZONDA, 
									double chastota_discret_pily, vector &vec_sokrashPILA)
{	
	time_t start_time, finish_time;
	time(&start_time);
	
	NumPointsOfOriginalPila = (NumPointsOfOriginalPila == 0) ?  chastota_discret_pily/chastota_pily : NumPointsOfOriginalPila;
	vector vecShumPrav, vecShumLev, vecTOK_ZONDA_posrednik(vecTOK_ZONDA.GetSize());
	int i = 0, j = 0, k = 0, one_segment_width = chastota_discret_zonda/chastota_pily, segments_amount = 0;
	double koeff_MaxMin = 1.5, SUM_Noise_Left = 0, SUM_Noise_Right = 0, SUM_Noise = 0, otn_SUM = 2;
	
	//////////////////СЛАБО ФИЛЬТРУЕМ ДЛЯ ТОГО ЧТОБЫ УБРАТЬ ВСПЛЕСКИ//////////////////
	//ocmath_savitsky_golay(vecTOK_ZONDA, vecTOK_ZONDA_posrednik, vecTOK_ZONDA.GetSize(), vecTOK_ZONDA.GetSize()/1000, -1, 1, 0);
	//////////////////СЛАБО ФИЛЬТРУЕМ ДЛЯ ТОГО ЧТОБЫ УБРАТЬ ВСПЛЕСКИ//////////////////
	
	////////////////////////////// находим все участки на пиле при помощи пиков //////////////////////////////
	uint nDataSize = vecPILA_ZONDA.GetSize();
	vector vxPeaks(nDataSize), vyPeaks(nDataSize), vxData(nDataSize);
	vector<int> vnIndices(nDataSize);
	
	if(max(vecPILA_ZONDA) > 0)
	{
		if( ocmath_find_peaks_by_local_maximum( &nDataSize, vxData, vecPILA_ZONDA, 
			vxPeaks, vyPeaks, vnIndices, POSITIVE_DIRECTION, (chastota_discret_pily/chastota_pily)*0.9) < OE_NOERROR ) return 0;
	}
	else
	{
		if( ocmath_find_peaks_by_local_maximum( &nDataSize, vxData, vecPILA_ZONDA * -1, 
			vxPeaks, vyPeaks, vnIndices, POSITIVE_DIRECTION, (chastota_discret_pily/chastota_pily)*0.9) < OE_NOERROR ) return 0;
	}
	if(nDataSize <= 1) return 0;
	vnIndices.SetSize(nDataSize);
	////////////////////////////// находим все участки на пиле при помощи пиков //////////////////////////////
	
	////////////////////////////// линейно аппроксимируем пилу и задаем ей размер с учетом otsech_nachala и otsech_konca //////////////////////////////
	vec_sokrashPILA = fit_linear_pila(vecPILA_ZONDA, vnIndices[0], one_segment_width);
	if(vec_sokrashPILA.GetSize()<=0) return 0;
	////////////////////////////// линейно аппроксимируем пилу и задаем ей размер с учетом otsech_nachala и otsech_konca //////////////////////////////
	
	////////////////////////////// приводим индексы начал отрезков к одной фазе с током //////////////////////////////
	if(chastota_discret_zonda != chastota_discret_pily) vnIndices = vnIndices/(chastota_discret_pily/chastota_discret_zonda);
	/* проверка на превышение размера вектора */
	if(vnIndices[vnIndices.GetSize()-1] > vecTOK_ZONDA.GetSize()-1) vnIndices = vnIndices*(chastota_discret_pily/chastota_discret_zonda);
	////////////////////////////// приводим индексы начал отрезков к одной фазе с током //////////////////////////////
	
	/* ВАЖНО!!!!!!!!!!!
	Мой метод предполагает, что либо в конце сигнала, либо в начале есть немного шума(чтобы взять его за образец и искать сигнал относительно него)
	*/
	vecShumLev.SetSize(one_segment_width);
	vecShumPrav.SetSize(one_segment_width);
	if(vecShumLev.GetSize()>vecTOK_ZONDA.GetSize()) return 0;
	
	////////////////////////////// заполняем шумовые вектора //////////////////////////////
	noise_vecs(one_segment_width, k, vecShumLev, vecShumPrav, vecTOK_ZONDA, SUM_Noise_Left, SUM_Noise_Right, vnIndices);
	////////////////////////////// заполняем шумовые вектора //////////////////////////////
	
	vector mnL_mxL_mnR_mxR(6);
	/*
	вектор mnL_mxL_mnR_mxR:
		-[0] минимум шума слева
		-[1] максимум шума справа
		-[2] минимум шума справа
		-[3] максимум шума справа
		-[4] минимум всего тока
		-[5] максимум всего тока
	*/
	
	////////////////////////////// заполняем вектор mnL_mxL_mnR_mxR //////////////////////////////
	MinMaxNoise(koeff_MaxMin, AverageSignal, vecShumLev, vecShumPrav, mnL_mxL_mnR_mxR);
	////////////////////////////// заполняем вектор mnL_mxL_mnR_mxR //////////////////////////////
	
	////////////////////////////// если пользователь сам выбрал диапазон обработки -> не зайдем в цикл while //////////////////////////////
	if(st_time_end_time[0] != -1 && st_time_end_time[1] != -1) convert_time_to_pts(vecTOK_ZONDA.GetSize(), chastota_discret_zonda);
	
	////////////////////////////// ОСНОВНОЙ ЦИКЛ ПОИСКА СИГНАЛА (ОЧИСТКИ ОТ ШУМА) //////////////////////////////
	while( segments_amount == 0 && st_time_end_time[0] == -1 && st_time_end_time[1] == -1)
	{
		////////// ДЛЯ ОТЛАДКИ //////////
		//Worksheet wks;
		//wks.Create("Origin");
		//Dataset ds1(wks.Columns(0)), ds2(wks.Columns(1));
		//ds1 = vecShumLev;
		//ds2 = vecShumPrav;
		
		k++;
		if(k > 5) return 0;
		
		i_nach = find_i_nach(vecTOK_ZONDA, SUM_Noise_Left, SUM_Noise_Right, one_segment_width, mnL_mxL_mnR_mxR, otn_SUM, vnIndices);
		i_konec = find_i_konec(vecTOK_ZONDA, SUM_Noise_Left, SUM_Noise_Right, one_segment_width, mnL_mxL_mnR_mxR, otn_SUM, vnIndices);
		
		if(i_nach < i_konec && i_nach != 0 && i_konec != 0) segments_amount = abs( i_konec - i_nach ) / one_segment_width;
		else return 0;
		
		/* если количество всех отрезков меньше 0.1 от потенциально максимального, то их слишком мало -> пересчитываем (нужно увеличить количество) */
		if( segments_amount <= (int)(0.1*vecTOK_ZONDA.GetSize()/(one_segment_width)) )
		{
			////////////////////////////// уменьшаем коэффициенты чтобы понизить пороги отбора отрезков //////////////////////////////
			koeff_MaxMin -= 0.3;
			otn_SUM *= 0.8;
			
			////////////////////////////// снова заполняем шумовые вектора //////////////////////////////
			noise_vecs(one_segment_width, k, vecShumLev, vecShumPrav, vecTOK_ZONDA, SUM_Noise_Left, SUM_Noise_Right, vnIndices);
			
			////////////////////////////// снова заполняем вектор mnL_mxL_mnR_mxR //////////////////////////////
			MinMaxNoise(koeff_MaxMin, AverageSignal, vecShumLev, vecShumPrav, mnL_mxL_mnR_mxR);
			
			////////////////////////////// зануляем колво отрезков и начала/конца чтобы снова зайти в цикл while //////////////////////////////
			segments_amount = 0;
			i_nach = 0; 
			i_konec = 0;
		}
		/* если количество всех отрезков юольше 0.9 от потенциально максимального, то их слишком много -> пересчитываем (нужно уменьшить количество) */
		if( segments_amount >= (int)(0.8*vecTOK_ZONDA.GetSize()/(one_segment_width)) )
		{
			////////////////////////////// увеличиваем коэффициенты чтобы повысить пороги отбора отрезков //////////////////////////////
			koeff_MaxMin += 0.3;
			otn_SUM /= 0.8;
			
			////////////////////////////// снова заполняем шумовые вектора //////////////////////////////
			noise_vecs(one_segment_width, k, vecShumLev, vecShumPrav, vecTOK_ZONDA, SUM_Noise_Left, SUM_Noise_Right, vnIndices);
			
			////////////////////////////// снова заполняем вектор mnL_mxL_mnR_mxR //////////////////////////////
			MinMaxNoise(koeff_MaxMin, AverageSignal, vecShumLev, vecShumPrav, mnL_mxL_mnR_mxR);
			
			////////////////////////////// зануляем колво отрезков и начала/конца чтобы снова зайти в цикл while //////////////////////////////
			segments_amount = 0;
			i_nach = 0; 
			i_konec = 0;
		}
	}
	////////////////////////////// ОСНОВНОЙ ЦИКЛ ПОИСКА СИГНАЛА (ОЧИСТКИ ОТ ШУМА) //////////////////////////////
	
	/* проверка на случай если найдено слишком много отрезков, например сигнал - шум, а не полезный ток */
	if( abs( i_konec - i_nach ) / one_segment_width >= (int)( 0.8*vecTOK_ZONDA.GetSize() / ( one_segment_width ) ) ) return 0;
	
	////////////////////////////// ВАЖНО! ЗАДАЕМ РАЗМЕРА ВЕКТОРА НАЧАЛ ОТРЕЗКОВ //////////////////////////////
	vec_Otr_Start_Indxs.SetSize( abs( i_konec - i_nach ) / one_segment_width );
	////////////////////////////// ВАЖНО! ЗАДАЕМ РАЗМЕРА ВЕКТОРА НАЧАЛ ОТРЕЗКОВ //////////////////////////////
	
	for(i = 0, j=0; j < vec_Otr_Start_Indxs.GetSize() && i < vnIndices.GetSize(); i++)
	{
		if(vnIndices[i] > i_nach && vnIndices[i] < i_konec) 
		{
			vec_Otr_Start_Indxs[j] = vnIndices[i];
			j++;
		}
	}
	for(i = vec_Otr_Start_Indxs.GetSize()-1; i > 1; i--) 
	{
		if(vec_Otr_Start_Indxs[i] != 0)
		{
			vec_Otr_Start_Indxs.SetSize(i+1);
			break;
		}
	}
	
	if(vec_Otr_Start_Indxs.GetSize()<=0) return 0;
	
	time(&finish_time);
	printf("Run time find otr %d secs\n", finish_time - start_time);
	
	return -1;
}

/////////////////// УСЛОВИЕ ОЧИСТКИ ОТ ШУМА ///////////////////
/*
Точка начала сигнала должна обладать значением, либо большим максимума всего шума (mnL_mxL_mnR_mxR[5]), либо меньшим минимума всего шума (mnL_mxL_mnR_mxR[4])
-- это нужно чтобы исключить вхождение по массе при очень маленькой амплитуде сигнала(почти шум)  --
&&
	Точка начала сигнала && последующая точка должны обладать значениями, меньшими минимума шума слева (mnL_mxL_mnR_mxR[0]) домноженного на коэффициент (koeff_MaxMin)
	-- это может сразу найти сигнал по амплитуде, например если шум мал, а сигнал высокой амплитуды --
	-- две точки подряд взяты для исключения случайного всплеска --
	||
	Точка начала сигнала && последующая точка должны обладать значениями, большим максимума шума слева (mnL_mxL_mnR_mxR[1]) домноженного на коэффициент (koeff_MaxMin)
	-- это может сразу найти сигнал по амплитуде, например если шум мал, а сигнал высокой амплитуды --
	-- две точки подряд взяты для исключения случайного всплеска --
	||
	Точка начала сигнала && последующая точка должны обладать значениями, меньшими минимума шума справа (mnL_mxL_mnR_mxR[2]) домноженного на коэффициент (koeff_MaxMin)
	-- это может сразу найти сигнал по амплитуде, например если шум мал, а сигнал высокой амплитуды --
	-- две точки подряд взяты для исключения случайного всплеска --
	||
	Точка начала сигнала && последующая точка должны обладать значениями, большим максимума шума справа (mnL_mxL_mnR_mxR[3]) домноженного на коэффициент (koeff_MaxMin)
	-- это может сразу найти сигнал по амплитуде, например если шум мал, а сигнал высокой амплитуды --
	-- две точки подряд взяты для исключения случайного всплеска --
	||
		Счетчик(counter) должен быть больше или равен длине одного промежутка
		-- это нужно для того, чтобы сравнивать шум(SUM_Noise) только в моменты целого промежутка, а не в каждой точке --
		-- счетчик(counter) накидывается в каждой точке и обнуляется при достижении значения длины промежутка --
		&&
			.......ЕСЛИ ДВИЖЕНИЕ СЛЕВА НАПРАВО.......
			Текущий участок шума (SUM_Noise), иными словами его вес(сумма амплитуд), должен быть в (otn_SUM) раза больше(тяжелее) участка шума слева
			-- это условие работает, так как у шума сигнал колеблется около нуля, и сумма его амплитуд мала, а у сигнала амплитуды большие, и они дадут большый "вес" --
			||
			Текущий участок шума (SUM_Noise), иными словами его вес(сумма амплитуд), должен быть в (otn_SUM) раза меньше(легче) участка шума слева
			||
			Текущий участок шума (SUM_Noise), иными словами его вес(сумма амплитуд), должен быть в (otn_SUM) раза больше(тяжелее) участка шума справа
			-- это условие работает, так как у шума сигнал колеблется около нуля, и сумма его амплитуд мала, а у сигнала амплитуды большие, и они дадут большый "вес" --
			
			-- условия в два раза легче чем справа нет, на случай если сигнал справа начинается впритык(шума справа нет) --
			.......ЕСЛИ ДВИЖЕНИЕ СЛЕВА НАПРАВО.......
			
			.......ЕСЛИ ДВИЖЕНИЕ СПРАВА НАЛЕВО.......
			Текущий участок шума (SUM_Noise), иными словами его вес(сумма амплитуд), должен быть в (otn_SUM) раза больше(тяжелее) участка шума слева
			-- это условие работает, так как у шума сигнал колеблется около нуля, и сумма его амплитуд мала, а у сигнала амплитуды большие, и они дадут большый "вес" --
			||
			Текущий участок шума (SUM_Noise), иными словами его вес(сумма амплитуд), должен быть в (otn_SUM) раза больше(тяжелее) участка шума справа
			-- это условие работает, так как у шума сигнал колеблется около нуля, и сумма его амплитуд мала, а у сигнала амплитуды большие, и они дадут большый "вес" --
			||
			Текущий участок шума (SUM_Noise), иными словами его вес(сумма амплитуд), должен быть в (otn_SUM) раза меньше(легче) участка шума справа
			
			-- условия в два раза легче чем слева нет, на случай если сигнал слева начинается впритык(шума слева нет) --
			.......ЕСЛИ ДВИЖЕНИЕ СПРАВА НАЛЕВО.......
*/
/////////////////// УСЛОВИЕ ОЧИСТКИ ОТ ШУМА ///////////////////

int find_i_nach(vector &vecTOK_ZONDA, float SUM_Noise_Left, float SUM_Noise_Right, int one_segment_width, vector &mnL_mxL_mnR_mxR, float otn_SUM, vector<int> &vnIndices)
{
	float SUM_Noise = 0, otsech = (otsech_nachala >= otsech_konca) ? otsech_nachala : otsech_konca;
	int nach_ignored_gap = 0, konec_ignored_gap = 0;
	
	nach_ignored_gap = vnIndices[0] + (1 - otsech)*one_segment_width;
	konec_ignored_gap = vnIndices[0] + (1 + otsech)*one_segment_width;
	
	/* 
	Шагаю не по 1, а по 2, так обработка в 2 раза быстрее(очевидно), а точность не теряется
	ТАКЖЕ ВАЖНО, ЧТО НАЧИНАЮ ИДТИ НЕМНОГО ОТСТУПЯ ОТ ВСПЛЕСКА (vnIndices[0] + one_segment_width*otsech)
	*/
	for(int i = vnIndices[0] + one_segment_width*otsech, counter = 0, k = 0; i < vnIndices[vnIndices.GetSize()-1]; i+=2)
	{
		if( (i < nach_ignored_gap) || (i > konec_ignored_gap) )
		{
			if(counter >= one_segment_width) counter = SUM_Noise = 0;
			counter+=2;
			
			if( i > konec_ignored_gap )
			{
				k++;
				nach_ignored_gap = vnIndices[k] + (1 - otsech)*one_segment_width;
				konec_ignored_gap = vnIndices[k] + (1 + otsech)*one_segment_width;
			}
			
			SUM_Noise += abs(vecTOK_ZONDA[i]) + abs(vecTOK_ZONDA[i-1]);
			
			/////////////////// ДЛЯ ОТЛАДКИ УСЛОВИЯ СЛЕВА НАПРАВО ///////////////////
			int a1=0,a2=0,a3=0,a4=0,a5=0,a6=0;
			//double currentA = vecTOK_ZONDA[i];
			//if(( vecTOK_ZONDA[i] < mnL_mxL_mnR_mxR[0] && vecTOK_ZONDA[i+1] < mnL_mxL_mnR_mxR[0] )) a1++;
			//if(( vecTOK_ZONDA[i] > mnL_mxL_mnR_mxR[1] && vecTOK_ZONDA[i+1] > mnL_mxL_mnR_mxR[1] )) a2++;
			//if(( vecTOK_ZONDA[i] < mnL_mxL_mnR_mxR[2] && vecTOK_ZONDA[i+1] < mnL_mxL_mnR_mxR[2] )) a3++;
			//if(( vecTOK_ZONDA[i] > mnL_mxL_mnR_mxR[3] && vecTOK_ZONDA[i+1] > mnL_mxL_mnR_mxR[3] )) a4++;
			if(( ( SUM_Noise/SUM_Noise_Left > 1*otn_SUM 
				|| SUM_Noise/SUM_Noise_Left < 1/otn_SUM 
				|| SUM_Noise/SUM_Noise_Right > 1*otn_SUM 
				/*|| SUM_Noise/SUM_Noise_Right < 1/otn_SUM*/ ) 
					&& counter >= one_segment_width )) a5++;
			//if(( vecTOK_ZONDA[i] < mnL_mxL_mnR_mxR[4] || vecTOK_ZONDA[i] > mnL_mxL_mnR_mxR[5] )) a6++;
			/////////////////// ДЛЯ ОТЛАДКИ УСЛОВИЯ СЛЕВА НАПРАВО ///////////////////
			
			if( ( vecTOK_ZONDA[i] < mnL_mxL_mnR_mxR[4] || vecTOK_ZONDA[i] > mnL_mxL_mnR_mxR[5] ) 
				&& ( ( vecTOK_ZONDA[i] < mnL_mxL_mnR_mxR[0] && vecTOK_ZONDA[i+1] < mnL_mxL_mnR_mxR[0] ) 
					|| ( vecTOK_ZONDA[i] > mnL_mxL_mnR_mxR[1] && vecTOK_ZONDA[i+1] > mnL_mxL_mnR_mxR[1] )
					|| ( vecTOK_ZONDA[i] < mnL_mxL_mnR_mxR[2] && vecTOK_ZONDA[i+1] < mnL_mxL_mnR_mxR[2] ) 
					|| ( vecTOK_ZONDA[i] > mnL_mxL_mnR_mxR[3] && vecTOK_ZONDA[i+1] > mnL_mxL_mnR_mxR[3] ) 
						|| ( ( SUM_Noise/SUM_Noise_Left > 1*otn_SUM 
							|| SUM_Noise/SUM_Noise_Left < 1/otn_SUM 
							|| SUM_Noise/SUM_Noise_Right > 1*otn_SUM 
							/*|| SUM_Noise/SUM_Noise_Right < 1/otn_SUM*/ ) 
								&& counter >= one_segment_width ) ) )
			{
				////////// ДЛЯ ОТЛАДКИ //////////
				if(a5>0) MessageBox(GetWindow(), "Слева вошли по сумме");
				
				counter = SUM_Noise = 0;
				return i;
			}
		}
	}
	return 0;
}

int find_i_konec(vector &vecTOK_ZONDA, float SUM_Noise_Left, float SUM_Noise_Right, int one_segment_width, vector &mnL_mxL_mnR_mxR, float otn_SUM, vector<int> &vnIndices)
{
	float SUM_Noise = 0, otsech = (otsech_nachala >= otsech_konca) ? otsech_nachala : otsech_konca;
	int nach_ignored_gap = 0, konec_ignored_gap = 0;
	
	nach_ignored_gap = vnIndices[vnIndices.GetSize()-1] - (1 - otsech)*one_segment_width;
	konec_ignored_gap = vnIndices[vnIndices.GetSize()-1] - (1 + otsech)*one_segment_width;
	
	/* 
	Шагаю не по 1, а по 2, так обработка в 2 раза быстрее(очевидно), а точность не теряется
	ТАКЖЕ ВАЖНО, ЧТО НАЧИНАЮ ИДТИ НЕМНОГО ОТСТУПЯ ОТ ВСПЛЕСКА (vnIndices[vnIndices.GetSize()-1] - one_segment_width*otsech)
	*/
	for(int i = vnIndices[vnIndices.GetSize()-1] - one_segment_width*otsech, counter = 0, k = 0; i > vnIndices[0]; i-=2)
	{	
		if( (i > nach_ignored_gap) || (i < konec_ignored_gap) )
		{
			if(counter >= one_segment_width) counter = SUM_Noise = 0;
			counter+=2;
			
			if( i < konec_ignored_gap ) 
			{
				k++;
				nach_ignored_gap = vnIndices[vnIndices.GetSize()-1-k] - (1 - otsech)*one_segment_width;
				konec_ignored_gap = vnIndices[vnIndices.GetSize()-1-k] - (1 + otsech)*one_segment_width;
			}
			
			SUM_Noise += abs(vecTOK_ZONDA[i]) + abs(vecTOK_ZONDA[i+1]);
			
			/////////////////// ДЛЯ ОТЛАДКИ УСЛОВИЯ СПРАВА НАЛЕВО ///////////////////
			int b1=0,b2=0,b3=0,b4=0,b5=0,b6=0;
			//double currentB = vecTOK_ZONDA[i];
			//if(( vecTOK_ZONDA[i] < mnL_mxL_mnR_mxR[0] && vecTOK_ZONDA[i-1] < mnL_mxL_mnR_mxR[0] )) b1++;
			//if(( vecTOK_ZONDA[i] > mnL_mxL_mnR_mxR[1] && vecTOK_ZONDA[i-1] > mnL_mxL_mnR_mxR[1] )) b2++;
			//if(( vecTOK_ZONDA[i] < mnL_mxL_mnR_mxR[2] && vecTOK_ZONDA[i-1] < mnL_mxL_mnR_mxR[2] )) b3++;
			//if(( vecTOK_ZONDA[i] > mnL_mxL_mnR_mxR[3] && vecTOK_ZONDA[i-1] > mnL_mxL_mnR_mxR[3] )) b4++;
			if(( ( SUM_Noise/SUM_Noise_Left > 1*otn_SUM 
				/*|| SUM_Noise/SUM_Noise_Left < 1/otn_SUM*/ 
				|| SUM_Noise/SUM_Noise_Right > 1*otn_SUM 
				|| SUM_Noise/SUM_Noise_Right < 1/otn_SUM ) 
					&& counter >= one_segment_width )) b5++;
			//if( ( vecTOK_ZONDA[i] < mnL_mxL_mnR_mxR[4] || vecTOK_ZONDA[i] > mnL_mxL_mnR_mxR[5] )) b6++;
			/////////////////// ДЛЯ ОТЛАДКИ УСЛОВИЯ СПРАВА НАЛЕВО ///////////////////
			
			if( ( vecTOK_ZONDA[i] < mnL_mxL_mnR_mxR[4] || vecTOK_ZONDA[i] > mnL_mxL_mnR_mxR[5] ) 
				&& ( ( vecTOK_ZONDA[i] < mnL_mxL_mnR_mxR[0] && vecTOK_ZONDA[i-1] < mnL_mxL_mnR_mxR[0] ) 
					|| ( vecTOK_ZONDA[i] > mnL_mxL_mnR_mxR[1] && vecTOK_ZONDA[i-1] > mnL_mxL_mnR_mxR[1] ) 
					|| ( vecTOK_ZONDA[i] < mnL_mxL_mnR_mxR[2] && vecTOK_ZONDA[i-1] < mnL_mxL_mnR_mxR[2] ) 
					|| ( vecTOK_ZONDA[i] > mnL_mxL_mnR_mxR[3] && vecTOK_ZONDA[i-1] > mnL_mxL_mnR_mxR[3] ) 
					|| ( ( SUM_Noise/SUM_Noise_Left > 1*otn_SUM 
						/*|| SUM_Noise/SUM_Noise_Left < 1/otn_SUM*/ 
						|| SUM_Noise/SUM_Noise_Right > 1*otn_SUM 
						|| SUM_Noise/SUM_Noise_Right < 1/otn_SUM ) 
							&& counter >= one_segment_width ) ) )
			{
				////////// ДЛЯ ОТЛАДКИ //////////
				if(b5>0) MessageBox(GetWindow(), "Справа вошли по сумме");
				
				counter = SUM_Noise = 0;
				return i;
			}
		}
	}
	return 0;
}

void noise_vecs(int one_segment_width, int k, vector &vecShumLev, vector &vecShumPrav, vector &vecTOK, double &SUM_Noise_Left, double &SUM_Noise_Right, vector<int> &vnIndices)
{
	float otsech = (otsech_nachala >= otsech_konca) ? otsech_nachala : otsech_konca;
	int i = 0, j = 0, nach_ignored_gap = 0, konec_ignored_gap = 0;
	
	i = vnIndices[k] - one_segment_width*(1-otsech);
	if(i < 0) i = 0;
	
	for(j = 0; j < vecShumLev.GetSize(); i+=2)
	{
		nach_ignored_gap = vnIndices[k] - otsech*one_segment_width;
		konec_ignored_gap = vnIndices[k] + otsech*one_segment_width;
		
		if( (i < nach_ignored_gap) || (i > konec_ignored_gap) )
		{
			if( i == konec_ignored_gap + 1 ) k++;
			
			vecShumLev[j] = vecTOK[i];
			vecShumLev[j+1] = vecTOK[i+1];
			SUM_Noise_Left +=  abs(vecTOK[i]) + abs(vecTOK[i+1]);
			j+=2;
		}
	}
	
	i = vnIndices[vnIndices.GetSize()-1-k] + one_segment_width*(1-otsech);
	if(i > vecTOK.GetSize()-1) i = vecTOK.GetSize()-1;
	
	for(j = 0; j < vecShumPrav.GetSize(); i-=2)
	{
		nach_ignored_gap = vnIndices[vnIndices.GetSize()-1-k] + otsech*one_segment_width;
		konec_ignored_gap = vnIndices[vnIndices.GetSize()-1-k] - otsech*one_segment_width;
		
		if( (i > nach_ignored_gap) || (i < konec_ignored_gap) )
		{
			if( i == konec_ignored_gap - 1 ) k++;
			
			vecShumPrav[j] = vecTOK[i];
			vecShumPrav[j+1] = vecTOK[i-1];
			SUM_Noise_Right += abs(vecTOK[i]) + abs(vecTOK[i-1]);
			j+=2;
		}
	}
	
	double avL, avR;
	vecShumLev.Sum(avL);
	vecShumPrav.Sum(avR);
	AverageSignal = (avL + avR)/(vecShumLev.GetSize() + vecShumPrav.GetSize());
}

void MinMaxNoise(float koeff_MaxMin, float AverageSignal, vector &vecShumLev, vector &vecShumPrav, vector &mnL_mxL_mnR_mxR)
{	
	if(AverageSignal != 0)
	{
		mnL_mxL_mnR_mxR[0] = /*min(vecShumLev)*/AverageSignal - abs(AverageSignal - min(vecShumLev))*koeff_MaxMin;
		mnL_mxL_mnR_mxR[1] = /*max(vecShumLev)*/AverageSignal + abs(max(vecShumLev) - AverageSignal)*koeff_MaxMin;
		mnL_mxL_mnR_mxR[2] = /*min(vecShumPrav)*/AverageSignal - abs(AverageSignal - min(vecShumPrav))*koeff_MaxMin;
		mnL_mxL_mnR_mxR[3] = /*max(vecShumPrav)*/AverageSignal + abs(max(vecShumPrav) - AverageSignal)*koeff_MaxMin;
	}
	else
	{
		if(min(vecShumLev) < 0) 
			mnL_mxL_mnR_mxR[0] = min(vecShumLev)*koeff_MaxMin;
		else
			mnL_mxL_mnR_mxR[0] = min(vecShumLev)/koeff_MaxMin;
		if(max(vecShumLev) > 0) 
			mnL_mxL_mnR_mxR[1] = max(vecShumLev)*koeff_MaxMin;
		else
			mnL_mxL_mnR_mxR[1] = max(vecShumLev)/koeff_MaxMin;
		if(min(vecShumPrav) < 0) 
			mnL_mxL_mnR_mxR[2] = min(vecShumPrav)*koeff_MaxMin;
		else
			mnL_mxL_mnR_mxR[2] = min(vecShumPrav)/koeff_MaxMin;
		if(max(vecShumPrav) > 0) 
			mnL_mxL_mnR_mxR[3] = max(vecShumPrav)*koeff_MaxMin;
		else
			mnL_mxL_mnR_mxR[3] = max(vecShumPrav)/koeff_MaxMin;
	}
	
	////////////////////////////////////// Определяем средний уровень шума //////////////////////////////////////
	if( max(vecShumLev) >  max(vecShumPrav) ) mnL_mxL_mnR_mxR[5] = AverageSignal + abs(max(vecShumLev) - AverageSignal);
	else mnL_mxL_mnR_mxR[5] = AverageSignal + abs(max(vecShumPrav) - AverageSignal);

	if( min(vecShumLev) <  min(vecShumPrav) ) mnL_mxL_mnR_mxR[4] = AverageSignal - abs(AverageSignal - min(vecShumLev));
	else mnL_mxL_mnR_mxR[4] = AverageSignal - abs(AverageSignal - min(vecShumPrav));
	////////////////////////////////////// Определяем средний уровень шума //////////////////////////////////////
}

vector fit_linear_pila(vector &vecPILA, int ind_st_of_pil, int one_segment_width)
{
	//float a = min(vecPILA), b = max(vecPILA);
	
	int i = 0, j = 0, ind_st, ind_end;
	ind_st = (otsech_nachala >= 0.2) ? ind_st_of_pil + otsech_nachala*NumPointsOfOriginalPila : ind_st_of_pil + 0.2*NumPointsOfOriginalPila;
	ind_end = (otsech_konca >= 0.2) ? ind_st_of_pil + (1-otsech_konca)*NumPointsOfOriginalPila : ind_st_of_pil + (1-0.2)*NumPointsOfOriginalPila;
	float stVP = vecPILA[ind_st], endVP = vecPILA[ind_end], delta;
	vector vec_sokrashPILA((one_segment_width)*(1-otsech_nachala-otsech_konca));
	
	/* определяем приращение в одной точке пилы от предыдущей */
	float segment_length = endVP - stVP, segment_points_amount = abs(ind_st-ind_end), transform_factor = (float)one_segment_width/NumPointsOfOriginalPila;
	delta = segment_length / (segment_points_amount * transform_factor);
	
	vec_sokrashPILA[0] = (otsech_nachala >= 0.2) ? stVP : vecPILA[ind_st_of_pil + otsech_nachala*NumPointsOfOriginalPila];
	for(i = 1; i < vec_sokrashPILA.GetSize(); i++) 
	{
		vec_sokrashPILA[i] = vec_sokrashPILA[i-1] + delta;
	}
	
	//float c = min(vec_sokrashPILA), d = max(vec_sokrashPILA);
	
	return vec_sokrashPILA;
}

void koeff_fit_linear_stepennaya(vector &vParams, vector &vec_Y, vector &vec_X)
{		
	vector vec_Y_stepennoi, vec_X_stepennoi;
	vec_Y_stepennoi = vec_Y;
	vec_X_stepennoi = vec_X;

	LROptions stLROptions;
	vec_Y_stepennoi.SetSize(ceil(vec_Y_stepennoi.GetSize()*chast_to_fitln));
	vec_X_stepennoi.SetSize(ceil(vec_X_stepennoi.GetSize()*chast_to_fitln));
	FitParameter psFitParameter[2];  // two parameters	 
	RegStats stRegStats;  // regression statistics
	RegANOVA stRegANOVA;  // anova statistics
	int nRet = ocmath_linear_fit(vec_X_stepennoi, vec_Y_stepennoi, vec_Y_stepennoi.GetSize(), psFitParameter, NULL, 0, &stLROptions, &stRegStats, &stRegANOVA);
	if (nRet != STATS_NO_ERROR)
	{
		out_str("!linear_fit");
	}					
	vParams[0] = psFitParameter[0].Value;
	vParams[1] = psFitParameter[1].Value;
}

void create_legend_zond(double A, double B, double C, double D, double Dens, string vremya_v_colonke, string str_SVCH_or_VCH, string str_He_or_Ar, string pgOrigName, string nomera_strok, Worksheet &wks_sheetLegend, int kolvo_col, Worksheet &wks_SevCols)
{
	Worksheet wks;
	wks.Create("Table", CREATE_HIDDEN); 
	WorksheetPage wksPage = wks.GetPage();
	wksPage.AddLayer("A+B*x");
	if(kolvo_col>1 && wks_SevCols != NULL)
	{
		wksPage.AddLayer("SeveralColumns");
		wks_SevCols = wksPage.Layers(2);
		while( wks_SevCols.DeleteCol(0) );
		wks_SevCols.SetSize(-1,4);
	}
	wks_sheetLegend = wksPage.Layers(1);
	while( wks_sheetLegend.DeleteCol(0) );
	wks_sheetLegend.SetSize(-1,4);
	wks_sheetLegend.Columns(1).SetLongName("A+B*x");
	wks_sheetLegend.Columns(3).SetLongName("Пересечение с OX");
	string str;
	str.Format("Участок времени %s", vremya_v_colonke);

	wks.Columns(0).SetLongName(str);
	wks.Columns(1).SetComments("Значения");
	wks.Columns(0).SetComments(pgOrigName); 
	wks.Columns(1).SetLongName("Строки "+ nomera_strok);
	
	if(kolvo_col>1)
	{
		wks.SetCell(0, 0, "Тип разряда");
		wks.SetCell(0, 1, str_SVCH_or_VCH);

		wks.SetCell(1, 0, "Рабочее тело");
		wks.SetCell(1, 1, str_He_or_Ar);
	}
	else
	{
		wks.SetCell(0, 0, "Коэффициент A");
		wks.SetCell(0, 1, A);

		wks.SetCell(1, 0, "Коэффициент B");
		wks.SetCell(1, 1, B);

		wks.SetCell(2, 0, "Коэффициент C");
		wks.SetCell(2, 1, C);

		wks.SetCell(3, 0, "Коэффициент D (Температура электронов, эВ)");
		wks.SetCell(3, 1, D);

		wks.SetCell(4, 0, "Плотность плазмы, см^-3");
		wks.SetCell(4, 1, Dens);
		
		wks.SetCell(5, 0, "Тип разряда");
		wks.SetCell(5, 1, str_SVCH_or_VCH);

		wks.SetCell(6, 0, "Рабочее тело");
		wks.SetCell(6, 1, str_He_or_Ar);
	}
	
	wks.AutoSize();

	//3. Add table as link to graph
	GraphLayer gl = Project.ActiveLayer();	
	GraphObject grTable = gl.CreateLinkTable(wksPage.GetName(), wks);
	wks.ConnectTo(gl);
}

void myFit(vector vec_Y_stepennoi, vector vec_X_stepennoi, vector &vParams, int Num_iter)
{
	//time_t start_time, finish_time;
	//time(&start_time);
	
	double A=vParams[0],B=vParams[1],C=vParams[2],D=vParams[3],vkid=5;
	int kolvo_polozh_tochek=0;
	
	while(Num_iter!=0)
	{
		kolvo_polozh_tochek=0;
		for(int i = vec_X_stepennoi.GetSize()*0.6; i<vec_X_stepennoi.GetSize(); i++)
		{
			if(fx(vec_X_stepennoi[i],A,B,C,D)<vec_Y_stepennoi[i])
			{
				kolvo_polozh_tochek++;
			}
		}
		if(abs(vec_X_stepennoi.GetSize()*0.4-kolvo_polozh_tochek*2)<vec_X_stepennoi.GetSize()*0.01)break;
		
		//if(min_pervyi_index==0||min_vtoroi_index==0) C = ((0.75*max(vec_Y_stepennoi)-A-B*vec_X_stepennoi[vec_X_stepennoi.GetSize()-1])/(exp(vec_X_stepennoi[vec_X_stepennoi.GetSize()-1]/D)));
		//else C = ((A-B*vec_X_stepennoi[peresech_X_index])/(exp(vec_X_stepennoi[peresech_X_index]/D)));
		
		C = ((0.75*max(vec_Y_stepennoi)-A-B*vec_X_stepennoi[vec_X_stepennoi.GetSize()-1])/(exp(vec_X_stepennoi[vec_X_stepennoi.GetSize()-1]/D)));
		
		if(vec_X_stepennoi.GetSize()*0.4-kolvo_polozh_tochek*2<0)
		{
			D=D+vkid;
		}
		else
		{
			D=D-vkid;
		}
		//if(vkid>0.1)
		vkid=vkid/2;
		if(D>25||D<5)break;
		Num_iter--;
	}
	vParams[2]=C;
	vParams[3]=D;
	
	//time(&finish_time);
	//printf("Run time myFit %d secs\n", finish_time - start_time);
}

void setoch()
{
	Worksheet wks = Project.ActiveLayer(), wks_slave;
    if(!wks) 
    {
    	MessageBox(GetWindow(), "wks!");
    	return;
    } 
    
    if(wks.GetPage().GetLongName().Match("Зонд*") || wks.GetPage().GetLongName().Match("Сетка*") 
    	|| wks.GetPage().GetLongName().Match("Цилиндр*") || wks.GetPage().GetLongName().Match("Магнит*"))
	{
		MessageBox(GetWindow(), "Не тот лист! Выберите лист с результатами эксперимента");
		return;
	}
    
	vector<string> vBooksNames, vErrsBooksNames;
	foreach(WorksheetPage pg in Project.WorksheetPages)
	{ 
		int BeginVal;
		is_str_numeric_integer(get_str_from_str_list(pg.GetLongName(), "_", 0), &BeginVal);
		if(BeginVal>10000)vBooksNames.Add(pg.GetLongName());
	}
	
	double chastota_pily = ch_pil,chastota_discret_pily,chastota_discret_SETOCH,Chast_to_Filtrate = 0.4, Chast_to_Filtrate_Diff = 0;
    int SVCH_or_VCH, He_or_Ar, True_or_False, Obr_Torec, Obr_Shtanga, Error=-1, UseSameName_or_Not_Shtanga, UseSameName_or_Not_PILA, UseSameName_or_Not_Torec, Poly_Order = 1, z_s_c = 1;
	TreeNode tn1;
	
    GETN_BOX( treeTest );
    
    GETN_BEGIN_BRANCH(Start, "Fitting Options")
    GETN_MULTI_COLS_BRANCH(2, DYNA_MULTI_COLS_COMAPCT) GETN_OPTION_BRANCH(GETNBRANCH_OPEN)
    
		GETN_CHECK(decayT3,"Обработка данных сеточного с торца", true)
		 GETN_NUM(decayT1, "Сопротивление на сеточном, Ом", r_s1) GETN_EDITOR_SIZE_USER_OPTION("#5")
		GETN_CHECK(decayT4,"Обработка данных сеточного со штанги", false)
		 GETN_NUM(decayT2, "Сопротивление на сеточном, Ом", r_s2) GETN_EDITOR_SIZE_USER_OPTION("#5")
		
	GETN_END_BRANCH(Start)
    
    //GETN_NUM(decayT2, "Частота пилы, Гц", ch_pil)
    
    GETN_BEGIN_BRANCH(Details2, "Задать частоты дискретизации(по умолчанию автоопределение)") GETN_CHECKBOX_BRANCH(0)
    //GETN_OPTION_BRANCH(GETNBRANCH_CLOSED)
	GETN_NUM(decayT8, "Частота дискретизации пилы, Гц", ch_dsc_p)
    GETN_NUM(decayT9, "Частота дискретизации сеточного, Гц", ch_dsc_t)
    GETN_END_BRANCH(Details2)
    
    GETN_BEGIN_BRANCH(Details3, "Задать начало и конец обработки(по умолчанию автоопределение)") GETN_CHECKBOX_BRANCH(0)
    //GETN_OPTION_BRANCH(GETNBRANCH_CLOSED)
    GETN_SLIDEREDIT(decayT1, "Время начала, с", 0.3, "0.0|5.0|50") 
    GETN_SLIDEREDIT(decayT2, "Время конца, с", 1.5, "0.0|5.0|50") 
    GETN_END_BRANCH(Details3)
    
    GETN_SLIDEREDIT(decayT14, "Часть точек, которые будут отсечены от начала", otsech_nachala, "0.01|0.49|48")
    GETN_SLIDEREDIT(decayT15, "Часть точек, которые будут отсечены от конца", otsech_konca, "0.01|0.49|48")

    GETN_SLIDEREDIT(decayT5, "Часть точек фильтрации", Chast_to_Filtrate, "0.01|0.99|98")
    
    GETN_BEGIN_BRANCH(DifFilt, "Сглаживание производной") GETN_CHECKBOX_BRANCH(1)
    GETN_OPTION_BRANCH(GETNBRANCH_OPEN)
    GETN_COMBO(decayT1, "Степень сглаживающей функции", Poly_Order, "1|2|3|4|5")
    GETN_SLIDEREDIT(decayT2, "Часть точек фильтрации", Chast_to_Filtrate/2, "0.01|0.99|98")   
    GETN_END_BRANCH(DifFilt)
    
    GETN_RADIO_INDEX_EX(decayT11, "Тип импульса", 2, "СВЧ|ВЧ|ИЦР")
    GETN_RADIO_INDEX_EX(decayT13, "Рабочее тело", 0, "He|Ar")
    GETN_BEGIN_BRANCH(Details, "Нужно обработать несколько книг") GETN_CHECKBOX_BRANCH(0)
    for(int i_1 = 0; i_1<vBooksNames.GetSize(); i_1++) 
    {
		GETN_CHECK(NodeName, vBooksNames[i_1], true)
    }
    GETN_END_BRANCH(Details)
    if(0==GetNBox( treeTest, "Обработка данных сеточного" )) return;
	
    //////////////////////////// СЧИТЫВАЕМ ДАННЫЕ ИЗ ДИАЛОГА ////////////////////////////
    treeTest.GetNode("Start").GetNode("decayT3").GetValue(Obr_Torec);
	treeTest.GetNode("Start").GetNode("decayT4").GetValue(Obr_Shtanga);
	treeTest.GetNode("Start").GetNode("decayT1").GetValue(r_s1);
	treeTest.GetNode("Start").GetNode("decayT2").GetValue(r_s2);
	//treeTest.GetNode("decayT2").GetValue(chastota_pily);
	treeTest.GetNode("decayT11").GetValue(SVCH_or_VCH);
	treeTest.GetNode("decayT13").GetValue(He_or_Ar);
	
	treeTest.GetNode("decayT14").GetValue(otsech_nachala);
	treeTest.GetNode("decayT15").GetValue(otsech_konca);
	otsech_nachala = (otsech_nachala == 0) ? 0.01 : otsech_nachala;
	otsech_konca = (otsech_konca == 0) ? 0.01 : otsech_konca;
	
	treeTest.GetNode("decayT5").GetValue(Chast_to_Filtrate);
	Chast_to_Filtrate = Chast_to_Filtrate/2;
	if(treeTest.DifFilt.Use)
	{
		treeTest.GetNode("DifFilt").GetNode("decayT1").GetValue(Poly_Order);
		treeTest.GetNode("DifFilt").GetNode("decayT2").GetValue(Chast_to_Filtrate_Diff);
		Chast_to_Filtrate_Diff = Chast_to_Filtrate_Diff/2;
	}
	if(Obr_Torec == 0 && Obr_Shtanga == 0) Obr_Torec = 1;
	if(treeTest.Details3.Use)
	{
		treeTest.GetNode("Details3").GetNode("decayT1").GetValue(st_time_end_time[0]);
		treeTest.GetNode("Details3").GetNode("decayT2").GetValue(st_time_end_time[1]);
		if(st_time_end_time[0] >= st_time_end_time[1])
		{
			MessageBox(GetWindow(), "Ошибка. Исправьте время начала и конца обработки. Вы выбрали от " + st_time_end_time[0] + " до " + st_time_end_time[1]);
			return;
		}
	}
	//////////////////////////// СЧИТЫВАЕМ ДАННЫЕ ИЗ ДИАЛОГА ////////////////////////////
	
	string ReportString_Shtanga, ReportString_Torec, ReportString_PILA;
	Column colPILA = ColIdxByName(wks, s_ps);
	if(!colPILA)
	{
		GETN_BOX( treeTest1 );
		GETN_STR(STR, "!Не найден столбец с именем " + "\"" + s_ps + "\"" + "!", "") GETN_INFO
		GETN_INTERACTIVE(interactive, "Выберите столбец с нужными данными", "[Book1]Sheet1!B")
		GETN_OPTION_INTERACTIVE_CONTROL(ICOPT_RESTRICT_TO_ONE_DATA) 
		if(treeTest.Details.Use)
		{
			GETN_RADIO_INDEX_EX(decayT1, "Использовать колонку с таким же именем для расчета остальных книг", 0, "Да|Нет")
		}
		if(0==GetNBox( treeTest1, "Ошибка!" )) return;
	
		treeTest1.GetNode("interactive").GetValue(ReportString_PILA);
		treeTest1.GetNode("decayT1").GetValue(UseSameName_or_Not_PILA);
		ReportString_PILA = get_str_from_str_list(ReportString_PILA, "\"", 1);
		colPILA = ColIdxByName(wks, ReportString_PILA);
		if(!colPILA)
		{
			MessageBox(GetWindow(), "Ошибка. Возможно, выбранная вами колонка пустая");
			return;
		}
		colPILA.SetLongName(s_ps);
		UseSameName_or_Not_PILA = UseSameName_or_Not_PILA + 2;
	}
	
	Column colTOK_SETOCH_Torec, colTOK_SETOCH_Shtanga; 
	if(Obr_Torec == 1)
	{
		colTOK_SETOCH_Torec = ColIdxByName(wks, s_tst);
		if(!colTOK_SETOCH_Torec)
		{
			GETN_BOX( treeTest1 );
			GETN_STR(STR, "!Не найден столбец с именем " + "\"" + s_tst + "\"" + "!", "") GETN_INFO
			GETN_INTERACTIVE(interactive, "Выберите столбец с нужными данными", "[Book1]Sheet1!B")
			GETN_OPTION_INTERACTIVE_CONTROL(ICOPT_RESTRICT_TO_ONE_DATA) 
			if(treeTest.Details.Use)
			{
				GETN_RADIO_INDEX_EX(decayT1, "Использовать колонку с таким же именем для расчета остальных книг", 0, "Да|Нет")
			}
			if(0==GetNBox( treeTest1, "Ошибка!" )) return;
			
			treeTest1.GetNode("interactive").GetValue(ReportString_Torec);
			treeTest1.GetNode("decayT1").GetValue(UseSameName_or_Not_Torec);
			ReportString_Torec = get_str_from_str_list(ReportString_Torec, "\"", 1);
			colTOK_SETOCH_Torec = ColIdxByName(wks, ReportString_Torec);
			if(!colTOK_SETOCH_Torec)
			{
				MessageBox(GetWindow(), "Ошибка. Возможно, выбранная вами колонка пустая");
				return;
			}
			colTOK_SETOCH_Torec.SetLongName(s_tst);
			UseSameName_or_Not_Torec = UseSameName_or_Not_Torec + 2;
		}
	}
	if(Obr_Shtanga == 1)
	{
		colTOK_SETOCH_Shtanga = ColIdxByName(wks, s_tssh);
		if(!colTOK_SETOCH_Shtanga)
		{
			GETN_BOX( treeTest1 );
			GETN_STR(STR, "!Не найден столбец с именем " + "\"" + s_tssh + "\"" + "!", "") GETN_INFO
			GETN_INTERACTIVE(interactive, "Выберите столбец с нужными данными", "[Book1]Sheet1!B")
			GETN_OPTION_INTERACTIVE_CONTROL(ICOPT_RESTRICT_TO_ONE_DATA) 
			if(treeTest.Details.Use) 
			{
				GETN_RADIO_INDEX_EX(decayT1, "Использовать колонку с таким же именем для расчета остальных книг", 0, "Да|Нет")
			}
			if(0==GetNBox( treeTest1, "Ошибка!" )) return;
			
			treeTest1.GetNode("interactive").GetValue(ReportString_Shtanga);
			treeTest1.GetNode("decayT1").GetValue(UseSameName_or_Not_Shtanga);
			ReportString_Shtanga = get_str_from_str_list(ReportString_Shtanga, "\"", 1);
			colTOK_SETOCH_Shtanga = ColIdxByName(wks, ReportString_Shtanga);
			if(!colTOK_SETOCH_Shtanga)
			{
				MessageBox(GetWindow(), "Ошибка. Возможно, выбранная вами колонка пустая");
				return;
			}
			colTOK_SETOCH_Shtanga.SetLongName(s_tssh);
			UseSameName_or_Not_Shtanga = UseSameName_or_Not_Shtanga + 2;
		}
	}
   	vector vecErrs;
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
				Column colTOK_SETOCH_Torec_slave, colTOK_SETOCH_Shtanga_slave, colPILA_slave;
				Page pg1 = Project.FindPage(vBooksNames[i_2]);
				if(pg1.Layers(0) != wks) 
				{
					wks_slave = pg1.Layers(0);
					if(UseSameName_or_Not_Shtanga == 2) 
					{
						colTOK_SETOCH_Shtanga_slave = ColIdxByName(wks_slave, ReportString_Shtanga);
						if(!colTOK_SETOCH_Shtanga_slave)
						{
							vecErrs.Add(222);
							vErrsBooksNames.Add("Сеточный штанга "+vBooksNames[i_2]+": ");
							continue;
						}
					}
					if(UseSameName_or_Not_Torec == 2) 
					{
						colTOK_SETOCH_Torec_slave = ColIdxByName(wks_slave, ReportString_Torec);
						if(!colTOK_SETOCH_Torec_slave)
						{
							vecErrs.Add(222);
							vErrsBooksNames.Add("Сеточный торец "+vBooksNames[i_2]+": ");
							continue;
						}
					}
					if(UseSameName_or_Not_PILA == 2) 
					{
						colPILA_slave = ColIdxByName(wks_slave, ReportString_PILA);
						if(!colPILA_slave)
						{
							vecErrs.Add(333);
							vErrsBooksNames.Add("Пила сеточного "+vBooksNames[i_2]+": ");
							continue;
						}
					}
				}
				else
				{
					wks_slave = wks;
					colTOK_SETOCH_Torec_slave = colTOK_SETOCH_Torec;
					colTOK_SETOCH_Shtanga_slave = colTOK_SETOCH_Shtanga;
					colPILA_slave = colPILA;
				}
				colTOK_SETOCH_Shtanga_slave.SetLongName(s_tssh);
				colTOK_SETOCH_Torec_slave.SetLongName(s_tst);
				colPILA_slave.SetLongName(s_ps);

				if(Obr_Torec == 1 && Error < 0) 
				{
					if(!treeTest.Details2.Use)
					{
						Error = auto_chdsc(chastota_discret_pily, chastota_discret_SETOCH, colPILA_slave, colTOK_SETOCH_Torec_slave, ch_pil, z_s_c, wks_slave, 0);
						if( Error != -1 ) 
						{
							vecErrs.Add(Error);
							if(Error != 333) vErrsBooksNames.Add("Сеточный штанга "+vBooksNames[i_2]+": ");
						}
						Error = -1;
					}
					else
					{
						treeTest.GetNode("Details2").GetNode("decayT8").GetValue(chastota_discret_pily);
						treeTest.GetNode("Details2").GetNode("decayT9").GetValue(chastota_discret_SETOCH);
					}
					
					Error = setoch_obrab(r_s1, chastota_pily, chastota_discret_pily, chastota_discret_SETOCH, SVCH_or_VCH, He_or_Ar, wks_slave, colTOK_SETOCH_Torec_slave, 0, colPILA_slave, Chast_to_Filtrate, Chast_to_Filtrate_Diff, Poly_Order);
					if(Error == 404) return;
					if(Error != -1) 
					{
						vecErrs.Add(Error);
						if(Error != 333) vErrsBooksNames.Add("Сеточный торец "+vBooksNames[i_2]+": ");
						else vErrsBooksNames.Add("Пила сеточного "+vBooksNames[i_2]+": ");
					}
					Error = -1;
				}
				if(Obr_Shtanga == 1 && Error < 0) 
				{
					if(!treeTest.Details2.Use)
					{
						Error = auto_chdsc(chastota_discret_pily, chastota_discret_SETOCH, colPILA_slave, colTOK_SETOCH_Shtanga_slave, ch_pil, z_s_c, wks_slave, 1);
						if( Error != -1 ) 
						{
							vecErrs.Add(Error);
							if(Error != 333) vErrsBooksNames.Add("Сеточный штанга "+vBooksNames[i_2]+": ");
						}
						Error = -1;
					}
					else
					{
						treeTest.GetNode("Details2").GetNode("decayT8").GetValue(chastota_discret_pily);
						treeTest.GetNode("Details2").GetNode("decayT9").GetValue(chastota_discret_SETOCH);
					}
				
					Error = setoch_obrab(r_s2, chastota_pily, chastota_discret_pily, chastota_discret_SETOCH, SVCH_or_VCH, He_or_Ar, wks_slave, colTOK_SETOCH_Shtanga_slave, 1, colPILA_slave, Chast_to_Filtrate, Chast_to_Filtrate_Diff, Poly_Order);
					if(Error == 404) return;
					if(Error != -1) 
					{
						vecErrs.Add(Error);
						if(Error != 333) vErrsBooksNames.Add("Сеточный штанга "+vBooksNames[i_2]+": ");
					}
					Error = -1;
				}
			}
		}
	}
	else 
	{
		if(Obr_Torec == 1) 
		{
			if(!treeTest.Details2.Use)
			{
				Error = auto_chdsc(chastota_discret_pily, chastota_discret_SETOCH, colPILA, colTOK_SETOCH_Torec, ch_pil, z_s_c, wks, 0);
				if( Error != -1 ) 
				{
					MessageBox(GetWindow(), FuncErrors(Error));
					return;
				}
				Error = -1;
			}
			else
			{
				treeTest.GetNode("Details2").GetNode("decayT8").GetValue(chastota_discret_pily);
				treeTest.GetNode("Details2").GetNode("decayT9").GetValue(chastota_discret_SETOCH);
			}
			
			Error = setoch_obrab(r_s1, chastota_pily, chastota_discret_pily, chastota_discret_SETOCH, SVCH_or_VCH, He_or_Ar, wks, colTOK_SETOCH_Torec, 0, colPILA, Chast_to_Filtrate, Chast_to_Filtrate_Diff, Poly_Order);
			if(Error == 404) return;
			if( Error != -1 ) 
			{
				vecErrs.Add(Error);
				if(Error != 333) vErrsBooksNames.Add("Сеточный торец "+get_str_from_str_list(wks.GetPage().GetLongName(), "_", 0)+": ");
				else vErrsBooksNames.Add("Пила сеточного "+get_str_from_str_list(wks.GetPage().GetLongName(), "_", 0)+": ");
			}
			Error = -1;
		}
		if(Obr_Shtanga == 1) 
		{
			if(!treeTest.Details2.Use)
			{
				Error = auto_chdsc(chastota_discret_pily, chastota_discret_SETOCH, colPILA, colTOK_SETOCH_Shtanga, ch_pil, z_s_c, wks, 1);
				if( Error != -1 ) 
				{
					MessageBox(GetWindow(), FuncErrors(Error));
					return;
				}
				Error = -1;
			}
			else
			{
				treeTest.GetNode("Details2").GetNode("decayT8").GetValue(chastota_discret_pily);
				treeTest.GetNode("Details2").GetNode("decayT9").GetValue(chastota_discret_SETOCH);
			}
			
			Error = setoch_obrab(r_s2, chastota_pily, chastota_discret_pily, chastota_discret_SETOCH, SVCH_or_VCH, He_or_Ar, wks, colTOK_SETOCH_Shtanga, 1, colPILA, Chast_to_Filtrate, Chast_to_Filtrate_Diff, Poly_Order);
			if(Error == 404) return;
			if( Error != -1 ) 
			{
				vecErrs.Add(Error);
				if(Error != 333) vErrsBooksNames.Add("Сеточный штанга "+get_str_from_str_list(wks.GetPage().GetLongName(), "_", 0)+": ");
				else vErrsBooksNames.Add("Пила сеточного "+get_str_from_str_list(wks.GetPage().GetLongName(), "_", 0)+": ");
			}
			Error = -1;
		}
	}
	if(vecErrs.GetSize()!=0)
	{
		string str = "Были ошибки в следующих книгах: \n";
		for(int aboba=0; aboba<vecErrs.GetSize(); aboba++) str = str + vErrsBooksNames[aboba] + FuncErrors(vecErrs[aboba]) + "\n";
		MessageBox(GetWindow(), str);
	}
	
	////////////////// возвращяю эту хрень к изначальным значениям иначе он ее запоминает и ломается //////////////////
	st_time_end_time[0] = -1;
	st_time_end_time[1] = -1;
	////////////////// возвращяю эту хрень к изначальным значениям иначе он ее запоминает и ломается //////////////////
}

string FuncErrors(int Error)
{
	switch(Error)
	{
		case 222: 
			return "Ошибка! Не найдена колонка с током или она пустая";
		case 333: 
			return "Ошибка! Не найдена колонка с пилой или она пустая";
		case 444: 
			return "Ошибка! Не найдено отрезков с пилой или током. Проверьте эти колонки, их названия, и правильность введенной частоты дискретизации";
		case 555: 
			return "Ошибка! Пересчет плотности уже применялся";
		case 666: 
			return "Ошибка! Не найдена колонка с интерферометром или она пустая";
		case 777: 
			return "Ошибка! Не получилось определить частоту дискретизации автоматически";
	}
	return "Error";
}

bool turn_tok_upwards(vector &vTOK, double resistance, int up_down)
{
	if(up_down == 0 && up_down != -1)
	{
		vTOK /= resistance;
		return false;
	}
	else if(up_down != -1)
	{
		vTOK /= -resistance;
		return true;
	}
	
	double min = min(vTOK), max = max(vTOK);
	double vsum, vavg;
	vTOK.Sum(vsum);
	vavg = vsum/vTOK.GetSize();
	
	if( abs(vavg - max) >= abs(vavg - min)) 
	{
		vTOK /= resistance;
		return false;
	}
	else
	{
		vTOK /= -resistance;
		return true;
	}
}

int setoch_obrab(double resistance, double chastota_pily, double chastota_discret_pily, double chastota_discret_SETOCH, int SVCH_or_VCH, int He_or_Ar, Worksheet wks, Column colTOK_SETOCH, int Torec_or_Shtanga, Column colPILA, double Chast_to_Filtrate, double Chast_to_Filtrate_Diff, int Poly_Order)
{
	time_t start_time, finish_time;
	time(&start_time);

	if(!colPILA) colPILA = ColIdxByName(wks, s_ps);	
	if(Torec_or_Shtanga == 0 && !colTOK_SETOCH) colTOK_SETOCH = ColIdxByName(wks, s_tst);
	if(Torec_or_Shtanga == 1 && !colTOK_SETOCH) colTOK_SETOCH = ColIdxByName(wks, s_tssh);
	if(!colTOK_SETOCH) return 222;
	if(!colPILA) return 333;
   	vector vecPILA_SETOCH(colPILA), vecTOK_SETOCH(colTOK_SETOCH), vec_sokrashPILA, vec_Otr_Start_Indxs(0);
    int kolvo_otrezkov=0;

	vecPILA_SETOCH = vecPILA_SETOCH * mnosj_sig;
   		
	if( find_otrezki_toka_i_obrab_pily(vec_Otr_Start_Indxs, vecTOK_SETOCH, chastota_discret_SETOCH, chastota_pily, vecPILA_SETOCH, chastota_discret_pily, vec_sokrashPILA) == 0 ) return 444;
	kolvo_otrezkov = vec_Otr_Start_Indxs.GetSize()+1;
	    
	vector vecOBRAB_TOKA, vec_X_stepennoi, vecTIME_synth, vecDIFF_TOKA, vec_maxVecDiff, vec_Temp, vecPeakVoltage, vecEnergy;
	vecOBRAB_TOKA.SetSize(vec_sokrashPILA.GetSize()*kolvo_otrezkov);
	vecDIFF_TOKA.SetSize(vec_sokrashPILA.GetSize()*kolvo_otrezkov);
	vec_maxVecDiff.SetSize(kolvo_otrezkov-1);
	vecPeakVoltage.SetSize(kolvo_otrezkov-1);
	vecEnergy.SetSize(kolvo_otrezkov-1);
	vec_Temp.SetSize(kolvo_otrezkov-1);
	vecTIME_synth.SetSize(kolvo_otrezkov-1);
	vec_X_stepennoi = vec_sokrashPILA;
	
	/* переворачиваем ток чтобы смотрел вверх(если нужно), и делим на сопротивление */
	if(false) turn_tok_upwards(vecTOK_SETOCH, resistance, -1);
	else vecTOK_SETOCH /= resistance;
	
	time_t start_time1, finish_time1;
	time(&start_time1);
	
	string s_progress, pgOrigName = get_str_from_str_list(wks.GetPage().GetLongName(), "_", 0);
	switch (Torec_or_Shtanga)
	{
		case 0:
			s_progress = "Обработка сеточного торец ";
			break;
		case 1:
			s_progress = "Обработка сеточного штанга ";
			break;
	}
	
	progressBox prgbBox(s_progress + wks.GetPage().GetLongName());
	prgbBox.SetRange(0, kolvo_otrezkov-1);
	
	for(int ii=0, lolo = 0, lolo_1 = 0; ii<kolvo_otrezkov-1; ii++)
	{
		if(!prgbBox.Set(ii)) return 404;
		
		vecTIME_synth[ii]=vec_Otr_Start_Indxs[ii]*(1/chastota_discret_SETOCH);
		
		make_one_segment(1, vec_Otr_Start_Indxs, vecTOK_SETOCH, chastota_discret_SETOCH, chastota_pily, vec_sokrashPILA, Chast_to_Filtrate, vecDIFF_TOKA, vec_maxVecDiff, 
						vec_Temp, vecPeakVoltage, vecEnergy, vecOBRAB_TOKA, Chast_to_Filtrate_Diff, Poly_Order, lolo, lolo_1, ii, -1);
	}
	
	time(&finish_time1);
	printf("Run time fitting %d secs\n", finish_time1 - start_time1);
	
	Worksheet wks_test, wks_test_1, wks_test_2, wks_test_3;
	Page pg;
	wks_test.Create("Origin");
	pg = wks_test.GetPage();
	wks.ConnectTo(wks_test);
	
	create_wks_z_s_c(1, pg, SVCH_or_VCH, He_or_Ar, kolvo_otrezkov, vec_sokrashPILA, Chast_to_Filtrate, resistance, chastota_discret_pily, chastota_discret_SETOCH, 
					vec_Otr_Start_Indxs, chastota_pily, Torec_or_Shtanga, Chast_to_Filtrate_Diff, Poly_Order, pgOrigName);
	
	wks_test = WksIdxByName(pg, "OrigData");
    wks_test_1 = WksIdxByName(pg, "Filtration");
    wks_test_2 = WksIdxByName(pg, "Differentiation");
    wks_test_3 = WksIdxByName(pg, "Parameters");
	
	Dataset ds_STOLBOV(wks_test.Columns(0)), dsTIME_synth(wks_test_3.Columns(0)), dsMaxValue(wks_test_3.Columns(1)), dsTemp(wks_test_3.Columns(2)), dsPeakVoltage(wks_test_3.Columns(3)), dsEnergy(wks_test_3.Columns(4));
	
	dsTIME_synth=vecTIME_synth;
	dsMaxValue=vec_maxVecDiff;
	dsTemp=vec_Temp;
	ds_STOLBOV=vec_sokrashPILA;
	ds_STOLBOV.Attach(wks_test_1.Columns(0));
	ds_STOLBOV=vec_sokrashPILA;
	ds_STOLBOV.Attach(wks_test_2.Columns(0));
	ds_STOLBOV=vec_sokrashPILA;
	dsPeakVoltage=vecPeakVoltage;
	//dsEnergy=vecEnergy;
	
	for(int i_2=1, k=0, k_2=0; i_2<kolvo_otrezkov; i_2++)
	{
		ds_STOLBOV.Attach(wks_test.Columns(i_2));
		ds_STOLBOV.SetSize(vec_sokrashPILA.GetSize());
	
		for(int i=0, k_1 = vec_Otr_Start_Indxs[i_2-1] + (chastota_discret_SETOCH/chastota_pily)*otsech_nachala; i<vec_sokrashPILA.GetSize(); i++, k_1++)
		{
			ds_STOLBOV[i]=vecTOK_SETOCH[k_1]
		}

		ds_STOLBOV.Attach(wks_test_1.Columns(i_2));
		ds_STOLBOV.SetSize(vec_sokrashPILA.GetSize());
		for(int j=0; j<vec_sokrashPILA.GetSize(); j++, k++)
		{
			ds_STOLBOV[j]=vecOBRAB_TOKA[k];
		}
		ds_STOLBOV.Attach(wks_test_2.Columns(i_2));
		ds_STOLBOV.SetSize(vec_sokrashPILA.GetSize());
		for(int j_1=0; j_1<vec_sokrashPILA.GetSize(); j_1++, k_2++)
		{
			ds_STOLBOV[j_1]=vecDIFF_TOKA[k_2];
		}
	}
	
	set_active_layer(wks_test_3);
	
	rezult_graphs(wks_test.GetPage(), 1);
	
	time(&finish_time);
	printf("Run time all prog %d secs\n", finish_time - start_time);
	
	return -1;
}

void graphSetoch()
{
	time_t start_time, finish_time;
	time(&start_time);
	
	Worksheet wks_test = Project.ActiveLayer();
    if(!wks_test)
    {
    	MessageBox(GetWindow(), "wks!");
    	return;
    }
    Page pg = wks_test.GetPage();
    string pgOrigName = pg.GetLongName();
    int SVCH_or_VCH=0, kolvo_col, first_c_i, last_c_i, r1, r2, comp_evo=-1, set_or_cil;
    vector<int> vROWS, vCOLS, vr2, vc2;
    
    if(pgOrigName.Match("Сетка*")) set_or_cil = 0;
    if(pgOrigName.Match("Цилиндр*")) set_or_cil = 1;
    if(pgOrigName.Match("Магнит*")) set_or_cil = 2;
    
    Worksheet wks_sheet1, wks_sheet2, wks_sheet3, wks_sheet4, wks_sheetLegend, wks_SevCols;
    wks_sheet1 = WksIdxByName(pg, "OrigData");
    wks_sheet2 = WksIdxByName(pg, "Filtration");
    if(set_or_cil == 0) wks_sheet3 = WksIdxByName(pg, "Differentiation");
    else wks_sheet3 = WksIdxByName(pg, "Fit");
    wks_sheet4 = WksIdxByName(pg, "Parameters");
    if(!wks_sheet1 || !wks_sheet3 || !wks_sheet4)
    {
    	MessageBox(GetWindow(), "Ошибка! Не найден лист с данными");
		return;
    }
    
	if(!wks_test.GetSelectedRange(vROWS, vCOLS, vr2, vc2))
	{
		if(wks_test == wks_sheet4) MessageBox(GetWindow(), "!Selection (выберите строку)");
		else MessageBox(GetWindow(), "!Selection (выберите колонку)");
		return;
	}
	if(!wks_test.GetSelection(first_c_i, last_c_i, r1, r2))
	{
		if(wks_test == wks_sheet4) MessageBox(GetWindow(), "!Selection (выберите строку)");
		else MessageBox(GetWindow(), "!Selection (выберите колонку)");
		return;
	} 
    
    if(wks_test == wks_sheet4)
    {
    	if(vROWS.GetSize() >= r2-r1+1 || (vROWS.GetSize() <= r2-r1+1 && vROWS.GetSize() > 1)) 
    	{
    		for(int i = 0, j; i < vROWS.GetSize(); i++) 
    		{
    			for(j = i+1; j < vROWS.GetSize(); j++)
    			{
					if(vROWS[i] == vROWS[j])
					{
						vROWS.RemoveAt(j);
						j--;
					}
    			}
    		}
    		kolvo_col = vROWS.GetSize();
    	}
    	else
    	{
    		kolvo_col = r2-r1+1;
    		vROWS.SetSize(kolvo_col);
    		vROWS[0] = r1;
    		for(int i = 1; i < vROWS.GetSize(); i++) vROWS[i] = vROWS[i-1] + 1;
    	}
    	for(int i = 1; i < vROWS.GetSize(); i++) if(vROWS[i] == vROWS[i-1]) vROWS.RemoveAt(i);
    }
    else
    {
    	if(first_c_i == 0) first_c_i = 1;
    	if(vCOLS[0] == 0) vCOLS[0] = 1;
    	if(vCOLS.GetSize() >= last_c_i-first_c_i+1 || (vCOLS.GetSize() <= last_c_i-first_c_i+1 && vCOLS.GetSize() > 1))
    	{
    		for(int i = 0, j; i < vCOLS.GetSize(); i++)
    		{
    			for(j = i+1; j < vCOLS.GetSize(); j++)
    			{
					if(vCOLS[i] == vCOLS[j]) 
					{
						vCOLS.RemoveAt(j);
						j--;
					}
    			}
    		}
    		kolvo_col = vCOLS.GetSize();
    	}
    	else 
    	{
    		kolvo_col = last_c_i-first_c_i+1;
    		vCOLS.SetSize(kolvo_col);
    		vCOLS[0] = first_c_i;
    		for(int i = 1; i < vCOLS.GetSize(); i++) vCOLS[i] = vCOLS[i-1] + 1;
    	}
    }

    if(kolvo_col>1)
	{
		GETN_BOX( treeTest ); 
		GETN_RADIO_INDEX_EX(decayT1, "", 0, "Сравнить выбранные отрезки|Построить эволюцию выбранных отрезков") 
		if(0==GetNBox( treeTest, "График нескольких отрезков" )) return;
		treeTest.GetNode("decayT1").GetValue(comp_evo);
	}

    string str_SVCH_or_VCH = get_str_from_str_list(wks_sheet1.Columns(0).GetUnits(), " ", 0);
    string str_He_or_Ar = get_str_from_str_list(wks_sheet1.Columns(0).GetUnits(), " ", 1);
    
	Dataset dsMaxValue(wks_sheet4.Columns(1)), dsTemp(wks_sheet4.Columns(2)), dsPILA(wks_sheet1.Columns(0)), dsSTOLBA, 
		dsPeakVoltage(wks_sheet4.Columns(3)), dsOD, dsFlt, dsDfn, dsSevT, dsTimeSynth(wks_sheet4.Columns(0));
	vector vec_maxVecDiff_Y, vec_maxVecDiff_X, vec_Polyvisota, vecY_lev, vecY_prav, vec_index_sred_lev, vec_index_sred_prav;	
	
	if(wks_test == wks_sheet4)
    {
		if(dsTimeSynth[vROWS[vROWS.GetSize()-1]] == 0)
		{
			MessageBox(GetWindow(), "Ошибка во времени! Последний отрезок обработан с ошибкой (время начала равно 0). "
				+ "Для решения проблемы удалите отрезки со временем начала 0, или пересчитайте импульс");
			return;
		}
	}
	else
	{
		if(dsTimeSynth[vCOLS[vCOLS.GetSize()-1]-1] == 0)
		{
			MessageBox(GetWindow(), "Ошибка во времени! Последний отрезок обработан с ошибкой (время начала равно 0). "
				+ "Для решения проблемы удалите отрезки со временем начала 0, или пересчитайте импульс");
			return;
		}
	}
	
	DataRange dr1,dr2,dr3,dr4,dr5,dr6,dr7;
	double time_step;
	string str, str1;
	int k, size_pily=dsPILA.GetSize(), k_end, num_col;
	if(comp_evo != 0)
	{
		if(wks_test==wks_sheet4)
		{
			if(kolvo_col>1)
			{
				int i;
				dsOD.Create(0);
				if(wks_sheet2) dsFlt.Create(0);
				dsDfn.Create(0);
				for(i = 0; i < vROWS.GetSize(); i++)
				{
					Dataset dsSlave;
					dsSlave.Attach(wks_sheet1.Columns(vROWS[i]+1));
					dsOD.Append(dsSlave, REDRAW_NONE);
					if(wks_sheet2)
					{
						dsSlave.Attach(wks_sheet2.Columns(vROWS[i]+1));
						if(dsSlave.GetUpperBound()<=0)
							dsFlt.SetSize(dsFlt.GetSize()+wks_sheet2.Columns(0).GetUpperBound());
						else
							dsFlt.Append(dsSlave, REDRAW_NONE);
					}
					dsSlave.Attach(wks_sheet3.Columns(vROWS[i]+1));
					dsDfn.Append(dsSlave, REDRAW_NONE);
				}
				
				dsSevT.Create(dsOD.GetSize());
				dsSevT[0] = dsTimeSynth[vROWS[0]];
				for(i = 1; i < dsSevT.GetSize(); i++) dsSevT[i] = dsSevT[i-1] + (dsTimeSynth[vROWS[vROWS.GetSize()-1]] + 
					(dsTimeSynth[vROWS[vROWS.GetSize()-1]]-dsTimeSynth[vROWS[vROWS.GetSize()-2]]) - dsTimeSynth[vROWS[0]])/dsSevT.GetSize();
			}
			else 
			{
				num_col = vROWS[0]+1;
				dr1.Add(wks_sheet1, 0, "X");
				if(wks_sheet2) dr2.Add(wks_sheet2, 0, "X");
				dr6.Add(wks_sheet3, 0, "X");
				dr1.Add(wks_sheet1, vROWS[0]+1, "Y");
				if(wks_sheet2) dr2.Add(wks_sheet2, vROWS[0]+1, "Y");
				dr6.Add(wks_sheet3, vROWS[0]+1, "Y");
				
				vec_maxVecDiff_Y.SetSize(2);
				vec_maxVecDiff_X.SetSize(2);
				vec_index_sred_lev.SetSize(2);
				vec_index_sred_prav.SetSize(2);
				
				vector v_whm_w_peak(10);
				/*
				вектор v_whm_w_peak:
					whm
					-[0] ширина на полувысоте
					-[1] Y на полувысоте
					-[2] X на полувысоте слева
					-[3] X на полувысоте справа
					W
					-[4] значение коэффициента W
					-[5] Y при W
					-[6] X при W слева
					-[7] X при W справа
					Peak
					-[8] Y пика
					-[9] X пика
				*/
				
				vec_maxVecDiff_Y[0] = vec_maxVecDiff_Y[1] = dsMaxValue[vROWS[0]];
				vec_maxVecDiff_X[0] = min(dsPILA);
				vec_maxVecDiff_X[1] = max(dsPILA);
				
				dsSTOLBA.Attach(wks_sheet3.Columns(vROWS[0]+1));
				
				vec_Polyvisota.SetSize(2);
				vecY_lev.SetSize(2);
				vecY_prav.SetSize(2);
				vecY_lev[0] = vecY_prav[0] = min(dsSTOLBA);
				vecY_lev[1] = vecY_prav[1] = max(dsSTOLBA);
				
				v_whm_w_peak = find_whm_Wparam_LRmidindxs(dsPILA, dsSTOLBA, set_or_cil, dsTemp[vROWS[0]], -1);
				
				if(set_or_cil == 0)
				{
					vec_Polyvisota[0] = vec_Polyvisota[1] = v_whm_w_peak[1];
					vec_index_sred_lev[0] = vec_index_sred_lev[1] = v_whm_w_peak[2];
					vec_index_sred_prav[0] = vec_index_sred_prav[1] = v_whm_w_peak[3];
				}
				else
				{
					vec_Polyvisota[0] = vec_Polyvisota[1] = v_whm_w_peak[5];
					vec_index_sred_lev[0] = vec_index_sred_lev[1] = v_whm_w_peak[6];
					vec_index_sred_prav[0] = vec_index_sred_prav[1] = v_whm_w_peak[7];
				}
			}
		}
		else
		{
			if(kolvo_col>1)
			{
				int i;
				dsOD.Create(0);
				if(wks_sheet2) dsFlt.Create(0);
				dsDfn.Create(0);
				for(i = 0; i < vCOLS.GetSize(); i++)
				{
					Dataset dsSlave;
					dsSlave.Attach(wks_sheet1.Columns(vCOLS[i]));
					dsOD.Append(dsSlave, REDRAW_NONE);
					if(wks_sheet2)
					{
						dsSlave.Attach(wks_sheet2.Columns(vCOLS[i]));
						if(dsSlave.GetUpperBound()<=0)
							dsFlt.SetSize(dsFlt.GetSize()+wks_sheet2.Columns(0).GetUpperBound());
						else
							dsFlt.Append(dsSlave, REDRAW_NONE);
					}
					dsSlave.Attach(wks_sheet3.Columns(vCOLS[i]));
					dsDfn.Append(dsSlave, REDRAW_NONE);
				}
				
				dsSevT.Create(dsOD.GetSize());
				dsSevT[0] = dsTimeSynth[vCOLS[0]-1];
				for(i = 1; i < dsSevT.GetSize(); i++) dsSevT[i] = dsSevT[i-1] + (dsTimeSynth[vCOLS[vCOLS.GetSize()-1]-1] + 
					(dsTimeSynth[vCOLS[vCOLS.GetSize()-1]-1]-dsTimeSynth[vCOLS[vCOLS.GetSize()-1]-2]) - dsTimeSynth[vCOLS[0]-1])/dsSevT.GetSize();
			}
			else 
			{
				num_col = vCOLS[0];
				dr1.Add(wks_sheet1, 0, "X");
				if(wks_sheet2) dr2.Add(wks_sheet2, 0, "X");
				dr6.Add(wks_sheet3, 0, "X");
				dr1.Add(wks_sheet1, vCOLS[0], "Y");
				if(wks_sheet2) dr2.Add(wks_sheet2, vCOLS[0], "Y");
				dr6.Add(wks_sheet3, vCOLS[0], "Y");
				
				vec_maxVecDiff_Y.SetSize(2);
				vec_maxVecDiff_X.SetSize(2);
				vec_index_sred_lev.SetSize(2);
				vec_index_sred_prav.SetSize(2);
				
				vector v_whm_w_peak(10);
				/*
				вектор v_whm_w_peak:
					whm
					-[0] ширина на полувысоте
					-[1] Y на полувысоте
					-[2] X на полувысоте слева
					-[3] X на полувысоте справа
					W
					-[4] значение коэффициента W
					-[5] Y при W
					-[6] X при W слева
					-[7] X при W справа
					Peak
					-[8] Y пика
					-[9] X пика
				*/
				
				vec_maxVecDiff_Y[0] = vec_maxVecDiff_Y[1] = dsMaxValue[vCOLS[0]-1];
				vec_maxVecDiff_X[0] = min(dsPILA);
				vec_maxVecDiff_X[1] = max(dsPILA);
				
				dsSTOLBA.Attach(wks_sheet3.Columns(vCOLS[0]));
				
				vec_Polyvisota.SetSize(2);
				vecY_lev.SetSize(2);
				vecY_prav.SetSize(2);
				vecY_lev[0] = vecY_prav[0] = min(dsSTOLBA);
				vecY_lev[1] = vecY_prav[1] = max(dsSTOLBA);
				
				v_whm_w_peak = find_whm_Wparam_LRmidindxs(dsPILA, dsSTOLBA, set_or_cil, dsTemp[vCOLS[0]-1], -1);
				
				if(set_or_cil == 0)
				{
					vec_Polyvisota[0] = vec_Polyvisota[1] = v_whm_w_peak[1];
					vec_index_sred_lev[0] = vec_index_sred_lev[1] = v_whm_w_peak[2];
					vec_index_sred_prav[0] = vec_index_sred_prav[1] = v_whm_w_peak[3];
				}
				else
				{
					vec_Polyvisota[0] = vec_Polyvisota[1] = v_whm_w_peak[5];
					vec_index_sred_lev[0] = vec_index_sred_lev[1] = v_whm_w_peak[6];
					vec_index_sred_prav[0] = vec_index_sred_prav[1] = v_whm_w_peak[7];
				}
			}
		}
	}
	
	if(wks_test==wks_sheet4)
	{
		if(kolvo_col>1)
		{
			str = get_str_from_str_list(wks_sheet1.Columns(min(vROWS)+1).GetLongName(), " ", 0) + " - " + get_str_from_str_list(wks_sheet1.Columns(max(vROWS)+1).GetLongName(), " ", 2);
			str1 = get_str_from_str_list(wks_sheet1.Columns(min(vROWS)+1).GetUnits(), " ", 0) + " - " + get_str_from_str_list(wks_sheet1.Columns(max(vROWS)+1).GetUnits(), " ", 2);
			k=vROWS[0];
			k_end=vROWS[vROWS.GetSize()-1];
		}
		else 
		{
			str = wks_sheet1.Columns(vROWS[0]+1).GetLongName();
			str1 = wks_sheet1.Columns(vROWS[0]+1).GetUnits();
			k=vROWS[0];
		}
	}
	else
	{
		if(kolvo_col>1)
		{
			str = get_str_from_str_list(wks_sheet1.Columns(min(vCOLS)).GetLongName(), " ", 0) + " - " + get_str_from_str_list(wks_sheet1.Columns(max(vCOLS)).GetLongName(), " ", 2);
			str1 = get_str_from_str_list(wks_sheet1.Columns(min(vCOLS)).GetUnits(), " ", 0) + " - " + get_str_from_str_list(wks_sheet1.Columns(max(vCOLS)).GetUnits(), " ", 2);
			k=vCOLS[0]-1;
			k_end=vCOLS[vCOLS.GetSize()-1]-1;
		}
		else 
		{
			str = wks_sheet1.Columns(vCOLS[0]).GetLongName();
			str1 = wks_sheet1.Columns(vCOLS[0]).GetUnits();
			k=vCOLS[0]-1;
		}
	}
    
	GraphPage gp;
	gp.Create("Origin");
	if(set_or_cil == 0) page_add_layer(gp, false, false, false, true, ADD_LAYER_INIT_SIZE_POS_MOVE_OFFSET, false, 0, LINK_STRAIGHT);
	GraphLayer gl1 = gp.Layers(0), gl2 = gp.Layers(1); 
	wks_sheet1.ConnectTo(gl1);
	//wks_sheet1.ConnectTo(gl2);
	Tree tr, tr1, tr2, tr3, tr6;
	tr = gl1.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
	tr.Root.Legend.Font.Size.nVal = 12;
	if(0 == gl1.UpdateThemeIDs(tr.Root) ) bool bRet = gl1.ApplyFormat(tr, true, true);
	
	Axis axesX = gl1.XAxis, axesY = gl1.YAxis ;	
	axesX.Additional.ZeroLine.nVal = 1; 
	axesY.Additional.ZeroLine.nVal = 1; 
    
	if(set_or_cil == 0)
	{
		if(kolvo_col>1) gp.SetLongName ( FindPageNameNumber ( get_str_from_str_list(pgOrigName, " ", 0) + " " + get_str_from_str_list(pgOrigName, " ", 4) + '(' + get_str_from_str_list(pgOrigName, " ", 5) + ')' + " Отрезки " + (k+1) + " - " + (k_end+1) ) );
		else gp.SetLongName ( FindPageNameNumber ( get_str_from_str_list(pgOrigName, " ", 0) + " " + get_str_from_str_list(pgOrigName, " ", 4) + '(' + get_str_from_str_list(pgOrigName, " ", 5) + ')' + " Отрезок " + (k+1) ) );
	}
	else
	{
		if(kolvo_col>1) gp.SetLongName ( FindPageNameNumber ( get_str_from_str_list(pgOrigName, " ", 0) + " " + get_str_from_str_list(pgOrigName, " ", 3) + '(' + get_str_from_str_list(pgOrigName, " ", 4) + ')' + " Отрезки " + (k+1) + " - " + (k_end+1) ) );
		else gp.SetLongName ( FindPageNameNumber ( get_str_from_str_list(pgOrigName, " ", 0) + " " + get_str_from_str_list(pgOrigName, " ", 3) + '(' + get_str_from_str_list(pgOrigName, " ", 4) + ')' + " Отрезок " + (k+1) ) );
	}
    create_legend_setoch(dsTemp[k], str, str_SVCH_or_VCH, str_He_or_Ar, pgOrigName, str1, wks_sheetLegend, dsPeakVoltage[k], kolvo_col, wks_SevCols);
    if(comp_evo == 1 && kolvo_col > 1)
    {
    	Dataset dsSlave;
    	dsSlave.Create(dsOD.GetSize());
    	dsSlave.Attach(wks_SevCols.Columns(0));
    	dsSlave = dsOD;
    	if(wks_sheet2)
		{
			dsSlave.Attach(wks_SevCols.Columns(1));
			dsSlave = dsFlt;
		}
    	dsSlave.Attach(wks_SevCols.Columns(2));
    	dsSlave = dsDfn;
    	dsSlave.Attach(wks_SevCols.Columns(3));
    	dsSlave = dsSevT;
    }
    
    Dataset ds_maxVecDiff_X, ds_maxVecDiff_Y, ds_Polyvisota, ds_index_sred_lev, ds_index_sred_prav, ds_Y_lev, ds_Y_prav; 
    DataPlot dp1,dp2,dp3,dp4,dp5,dp6,dp7;

    if(kolvo_col == 1)
    {
		ds_maxVecDiff_X.Attach(wks_sheetLegend.Columns(0));
		ds_maxVecDiff_Y.Attach(wks_sheetLegend.Columns(1));
		ds_Polyvisota.Attach(wks_sheetLegend.Columns(2));
		ds_index_sred_lev.Attach(wks_sheetLegend.Columns(3));
		ds_Y_lev.Attach(wks_sheetLegend.Columns(4));
		ds_index_sred_prav.Attach(wks_sheetLegend.Columns(5));
		ds_Y_prav.Attach(wks_sheetLegend.Columns(6));
		
		ds_maxVecDiff_X=vec_maxVecDiff_X;
		ds_maxVecDiff_Y=vec_maxVecDiff_Y;
		ds_Polyvisota=vec_Polyvisota;
		ds_index_sred_lev=vec_index_sred_lev;
		ds_Y_lev=vecY_lev;
		ds_index_sred_prav=vec_index_sred_prav;
		ds_Y_prav=vecY_prav;
		
		dr3.Add(wks_sheetLegend, 0, "X");
		dr3.Add(wks_sheetLegend, 1, "Y");
		dr7.Add(wks_sheetLegend, 0, "X");
		dr7.Add(wks_sheetLegend, 2, "Y");
		dr4.Add(wks_sheetLegend, 3, "X");
		dr4.Add(wks_sheetLegend, 4, "Y");
		dr5.Add(wks_sheetLegend, 5, "X");
		dr5.Add(wks_sheetLegend, 6, "Y");
		
		gl1.AddPlot(dr1, IDM_PLOT_LINE);
		if(wks_sheet2 && wks_sheet2.Columns(num_col).GetUpperBound() > 0) gl1.AddPlot(dr2, IDM_PLOT_LINE);
		if(set_or_cil == 0)
		{
			gl2.AddPlot(dr3, IDM_PLOT_LINE);
			gl2.AddPlot(dr7, IDM_PLOT_LINE);
			gl2.AddPlot(dr4, IDM_PLOT_LINE);
			gl2.AddPlot(dr5, IDM_PLOT_LINE);
			gl2.AddPlot(dr6, IDM_PLOT_LINE);
		}
		else
		{
			gl1.AddPlot(dr3, IDM_PLOT_LINE);
			gl1.AddPlot(dr7, IDM_PLOT_LINE);
			gl1.AddPlot(dr4, IDM_PLOT_LINE);
			gl1.AddPlot(dr5, IDM_PLOT_LINE);
			gl1.AddPlot(dr6, IDM_PLOT_LINE);
		}
		
		if(set_or_cil == 0)
		{
			dp1 = gl1.DataPlots(0);
			dp2 = gl1.DataPlots(1);
			dp3 = gl2.DataPlots(0);
			dp7 = gl2.DataPlots(1);
			dp4 = gl2.DataPlots(2);
			dp5 = gl2.DataPlots(3);
			dp6 = gl2.DataPlots(4);
		}
		else
		{
			dp1 = gl1.DataPlots(0);
			if(wks_sheet2 && wks_sheet2.Columns(num_col).GetUpperBound() > 0)
			{
				dp2 = gl1.DataPlots(6);
				dp3 = gl1.DataPlots(2);
				dp7 = gl1.DataPlots(3);
				dp4 = gl1.DataPlots(4);
				dp5 = gl1.DataPlots(5);
				dp6 = gl1.DataPlots(1);
			}
			else
			{
				dp2 = gl1.DataPlots(5);
				dp3 = gl1.DataPlots(1);
				dp7 = gl1.DataPlots(2);
				dp4 = gl1.DataPlots(3);
				dp5 = gl1.DataPlots(4);
			}
		}
		
		tr1 = dp1.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
		tr2 = dp2.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
		tr3 = dp3.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
		if(wks_sheet2 && wks_sheet2.Columns(num_col).GetUpperBound() > 0) tr6 = dp6.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
		
		tr1.Root.Line.Width.nVal = 2;
		tr1.Root.Line.Color.nVal = set_color_rnd();
		
		tr2.Root.Line.Width.dVal=2.5;
		tr2.Root.Line.Color.nVal=1;
		
		tr3.Root.Line.Width.dVal=1.1;
		tr3.Root.Line.Color.nVal=3;

		tr6.Root.Line.Width.dVal=2.5;
		tr6.Root.Line.Color.nVal=2;
		
		dp1.ApplyFormat(tr1, true, true);
		dp2.ApplyFormat(tr2, true, true);
		dp3.ApplyFormat(tr3, true, true);
		dp4.ApplyFormat(tr3, true, true);
		dp5.ApplyFormat(tr3, true, true);
		if(wks_sheet2 && wks_sheet2.Columns(num_col).GetUpperBound() > 0) dp6.ApplyFormat(tr6, true, true);
		dp7.ApplyFormat(tr3, true, true);
    }
    if(comp_evo == 0 && kolvo_col > 1)
    {
    	Curve cc;
    	int i, k = 0, k2 = 0, color_l1 = 0, color_l2 = 0;
    	if(wks_test == wks_sheet4)
    	{
    		/////////////ORIGDATA/////////////
    		for(i = 0; i < vROWS.GetSize(); i++)
    		{
				cc.Attach(wks_sheet1, 0, vROWS[i]+1);
				gl1.AddPlot(cc, IDM_PLOT_SCATTER);
				tr1 = gl1.DataPlots(k).GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
				tr1.Root.Symbol.Size.nVal=2;
				gl1.DataPlots(k).ApplyFormat(tr1, true, true);
				color_l1 = set_color_except_bad(color_l1);
				gl1.DataPlots(k).SetColor(color_l1);//всего цветов 22
				k++;
    		}
    		/////////////SMOOTH/////////////
    		/*
			for(i = 0; i < vROWS.GetSize(); i++)
			{	
				if(wks_sheet4 && wks_sheet4.Columns(vROWS[i]+1).GetUpperBound() > 0)
				{
					cc.Attach(wks_sheet4, 0, vROWS[i]+1);
					gl1.AddPlot(cc, IDM_PLOT_LINE);
					tr1 = gl1.DataPlots(k).GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
					tr1.Root.Line.Width.dVal=2;
					gl1.DataPlots(k).ApplyFormat(tr1, true, true);
					gl1.DataPlots(k).SetColor(k+1);
					k++;
				}
			}
			*/
			/////////////FITTING/////////////
			if(set_or_cil == 1 || set_or_cil == 2)
			{
				for(i = 0; i < vROWS.GetSize(); i++)
				{
					cc.Attach(wks_sheet3, 0, vROWS[i]+1);
					gl1.AddPlot(cc, IDM_PLOT_LINE);
					tr1 = gl1.DataPlots(k).GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
					tr1.Root.Line.Width.dVal=2;
					gl1.DataPlots(k).ApplyFormat(tr1, true, true);
					color_l1 = set_color_except_bad(color_l1);
					gl1.DataPlots(k).SetColor(color_l1);
					k++;
				}
			}
			else
			{
				for(i = 0; i < vROWS.GetSize(); i++)
				{
					cc.Attach(wks_sheet3, 0, vROWS[i]+1);
					gl2.AddPlot(cc, IDM_PLOT_LINE);
					tr1 = gl2.DataPlots(k2).GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
					tr1.Root.Line.Width.dVal=2;
					gl2.DataPlots(k2).ApplyFormat(tr1, true, true);
					color_l2 = set_color_except_bad(color_l2);
					gl2.DataPlots(k2).SetColor(color_l2);
					k2++;
				}
			}
			
    	}
    	else
    	{
    		/////////////ORIGDATA/////////////
    		for(i = 0; i < vCOLS.GetSize(); i++)
    		{
				cc.Attach(wks_sheet1, 0, vCOLS[i]);
				gl1.AddPlot(cc, IDM_PLOT_SCATTER);
				tr1 = gl1.DataPlots(k).GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
				tr1.Root.Symbol.Size.nVal=2;
				gl1.DataPlots(k).ApplyFormat(tr1, true, true);
				color_l1 = set_color_except_bad(color_l1);
				gl1.DataPlots(k).SetColor(color_l1);
				k++;
    		}
    		/////////////SMOOTH/////////////
    		/*
			for(i = 0; i < vCOLS.GetSize(); i++)
			{
				if(wks_sheet4 && wks_sheet4.Columns(vCOLS[i]).GetUpperBound() > 0)
				{
					cc.Attach(wks_sheet4, 0, vCOLS[i]);
					gl1.AddPlot(cc, IDM_PLOT_LINE);
					tr1 = gl1.DataPlots(k).GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
					tr1.Root.Line.Width.dVal=2;
					gl1.DataPlots(k).ApplyFormat(tr1, true, true);
					gl1.DataPlots(k).SetColor(k+1);
					k++;
				}
			}
			*/
			/////////////FITTING/////////////
			if(set_or_cil == 1 || set_or_cil == 2)
			{
				for(i = 0; i < vCOLS.GetSize(); i++)
				{
					cc.Attach(wks_sheet3, 0, vCOLS[i]);
					gl1.AddPlot(cc, IDM_PLOT_LINE);
					tr1 = gl1.DataPlots(k).GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
					tr1.Root.Line.Width.dVal=2;
					gl1.DataPlots(k).ApplyFormat(tr1, true, true);
					color_l1 = set_color_except_bad(color_l1);
					gl1.DataPlots(k).SetColor(color_l1);
					k++;
				}
			}
			else
			{
				for(i = 0; i < vCOLS.GetSize(); i++)
				{
					cc.Attach(wks_sheet3, 0, vCOLS[i]);
					gl2.AddPlot(cc, IDM_PLOT_LINE);
					tr1 = gl2.DataPlots(k2).GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
					tr1.Root.Line.Width.dVal=2;
					gl2.DataPlots(k2).ApplyFormat(tr1, true, true);
					color_l2 = set_color_except_bad(color_l2);
					gl2.DataPlots(k2).SetColor(color_l2);
					k2++;
				}
			}
    	}
    }
    if(comp_evo == 1 && kolvo_col > 1)
    {
    	dr1.Add(wks_SevCols, 0, "Y");
    	dr1.Add(wks_SevCols, 3, "X");
		if(wks_sheet2)
		{
			dr2.Add(wks_SevCols, 1, "Y");
			dr2.Add(wks_SevCols, 3, "X");
		}
		dr6.Add(wks_SevCols, 2, "Y");
		dr6.Add(wks_SevCols, 3, "X");
		
		gl1.AddPlot(dr1, IDM_PLOT_LINE);
		if(wks_sheet2) gl1.AddPlot(dr2, IDM_PLOT_SCATTER);
		if(set_or_cil == 0) gl2.AddPlot(dr6, IDM_PLOT_SCATTER);
		else gl1.AddPlot(dr6, IDM_PLOT_SCATTER);
		
		dp1 = gl1.DataPlots(0);
		if(wks_sheet2) 
		{
			dp2 = gl1.DataPlots(1); 
			if(set_or_cil == 0) dp3 = gl2.DataPlots(0);
			else dp3 = gl1.DataPlots(2);
		}
		else
		{
			if(set_or_cil == 0) dp3 = gl2.DataPlots(0);
			else dp3 = gl1.DataPlots(1);
		}
		
		tr1 = dp1.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
		if(wks_sheet2) tr2 = dp2.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
		tr3 = dp3.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
		
		tr1.Root.Line.Width.dVal=1;
		tr1.Root.Line.Color.nVal=0;
		
		tr2.Root.Symbol.Size.nVal=2;
		tr2.Root.Symbol.EdgeColor.nVal=2;
		
		tr3.Root.Symbol.Size.nVal=2;
		tr3.Root.Symbol.EdgeColor.nVal=1;
		
		dp1.ApplyFormat(tr1, true, true);
		if(wks_sheet2) dp2.ApplyFormat(tr2, true, true);
		dp3.ApplyFormat(tr3, true, true);
    }
    
	Axis axisY = gl1.YAxis, axisX = gl1.XAxis;
	tr.Reset();
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
	
    if(comp_evo == 1) gl1.GraphObjects("XB").Text = "Время, с";
	else gl1.GraphObjects("XB").Text = "Напряжение, В";
	gl1.GraphObjects("YL").Text = "Ток, А";
	if(set_or_cil == 0) gl2.GraphObjects("YR").Text = "Аппаратная функция";
	
	gl1.Rescale();
	if(set_or_cil == 0) gl2.Rescale();
	legend_update(gl1, ALM_LONGNAME);
	legend_update(gl2, ALM_LONGNAME);
	legend_combine(gp);
	
	
	if(comp_evo == 0 || kolvo_col == 1)
    {
		if(dsPILA[0] >= 0) axisX.Scale.From.dVal = 0;
		else axisX.Scale.From.dVal = dsPILA[0];
		axisX.Scale.To.dVal = dsPILA[dsPILA.GetSize()-1];
		axisX.Scale.IncrementBy.nVal = 0; // 0=increment by value; 1=number of major ticks
		axisX.Scale.Value.dVal = 20; // Increment value
		
		axisY.Scale.IncrementBy.nVal = 1; // 0=increment by value; 1=number of major ticks
		axisY.Scale.MajorTicksCount.dVal = 10;
		if(set_or_cil == 0) axisY = gl2.YAxis;
		axisY.Scale.IncrementBy.nVal = 1; // 0=increment by value; 1=number of major ticks
		axisY.Scale.MajorTicksCount.dVal = 10;
		
		if(kolvo_col == 1)
		{
			if(set_or_cil == 0) create_labels(1);
			else create_labels(2);
		}
    }
    else 
    {
    	axisX.Scale.IncrementBy.nVal = 1; // 0=increment by value; 1=number of major ticks
		axisX.Scale.MajorTicksCount.dVal = 10;
		axisY.Scale.IncrementBy.nVal = 1; // 0=increment by value; 1=number of major ticks
		axisY.Scale.MajorTicksCount.dVal = 10;
    }
	
	time(&finish_time);
	printf("Run time graph %d secs\n", finish_time - start_time);
}


void create_legend_setoch(double Temp, string vremya_v_colonke, string str_SVCH_or_VCH, string str_He_or_Ar, string pgOrigName, string nomera_strok, Worksheet &wks_sheetLegend, double PeakVoltage, int kolvo_col, Worksheet &wks_SevCols)
{
	Worksheet wks;
	wks.Create("Table", CREATE_HIDDEN); 
	WorksheetPage wksPage = wks.GetPage();
	wksPage.AddLayer("GraphParams");
	if(kolvo_col>1)
	{
		wksPage.AddLayer("SeveralColumns");
		wks_SevCols = wksPage.Layers(2);
		while( wks_SevCols.DeleteCol(0) );
		wks_SevCols.SetSize(-1,4);
	}
	wks_sheetLegend = wksPage.Layers(1);
	while( wks_sheetLegend.DeleteCol(0) );
	wks_sheetLegend.SetSize(-1,7);
	wks_sheetLegend.Columns(1).SetLongName("MaxValue");
	wks_sheetLegend.Columns(2).SetLongName("HalfHigh");
	wks_sheetLegend.Columns(4).SetLongName("RightMid");
	wks_sheetLegend.Columns(6).SetLongName("LeftMid");

	string str;
	str.Format("Участок времени %s", vremya_v_colonke);	

	wks.Columns(0).SetLongName(str);
	wks.Columns(1).SetComments("Значения");
	wks.Columns(0).SetComments(pgOrigName); 
	wks.Columns(1).SetLongName("Строки "+ nomera_strok);

	if(kolvo_col>1)
	{
		wks.SetCell(0, 0, "Тип разряда");
		wks.SetCell(0, 1, str_SVCH_or_VCH);

		wks.SetCell(1, 0, "Рабочее тело");
		wks.SetCell(1, 1, str_He_or_Ar);
	}
	else
	{
		//wks.SetSize(4, 2);
		wks.SetCell(0, 0, "Температура");
		wks.SetCell(0, 1, Temp);

		wks.SetCell(1, 0, "Потенциал пика, В");
		wks.SetCell(1, 1, PeakVoltage);
		
		wks.SetCell(2, 0, "Тип разряда");
		wks.SetCell(2, 1, str_SVCH_or_VCH);

		wks.SetCell(3, 0, "Рабочее тело");
		wks.SetCell(3, 1, str_He_or_Ar);
	}
	
	wks.AutoSize();

	GraphLayer gl = Project.ActiveLayer();	
	GraphObject grTable = gl.CreateLinkTable(wksPage.GetName(), wks);
	wks.ConnectTo(gl);
}

int pereschet_otrezka_setoch(int &num_col, Worksheet &wks_graph_parent, Worksheet &wks_Legend)
{
	double Chast_to_Filtrate = 0.01, Chast_to_Filtrate_Diff;
	int index_nach_otrezka, index_konca_otrezka, size_otrezka, resistance, Poly_Order = 1, Filtered_or_Not = 0, set_or_cil, Filtered_or_Not_1 = 0;
	Worksheet wks_OrigData, wks_sheet1, wks_sheet2, wks_sheet3, wks_sheet4, wks_sheetLegend;
	
	if( Get_OriginData_wks(wks_graph_parent, wks_OrigData, wks_Legend) == 0 ) return 0;
	if(wks_Legend.Columns(0).GetUpperBound() < 3)
	{
		MessageBox(GetWindow(), "Невозможно пересчитать, на графике больше одного отрезка");
		return 0;
	}
	
	Page pg =  wks_graph_parent.GetPage(), pgLegend =  wks_Legend.GetPage();
	wks_sheetLegend = pgLegend.Layers(1);
	if(pg.GetLongName().Match("Сетка*")) set_or_cil = 0;
    if(pg.GetLongName().Match("Цилиндр*")||pg.GetLongName().Match("Магнит*")) set_or_cil = 1;
	
	is_str_numeric_integer(get_str_from_str_list(wks_graph_parent.Columns(0).GetUnits(), " ", 2), &resistance);
	is_str_numeric_integer(get_str_from_str_list(wks_Legend.Columns(1).GetLongName(), " ", 1), &index_nach_otrezka);
	is_str_numeric_integer(get_str_from_str_list(wks_Legend.Columns(1).GetLongName(), " ", 3), &index_konca_otrezka);
	size_otrezka = index_konca_otrezka - index_nach_otrezka + 1;
	
    wks_sheet1 = WksIdxByName(pg, "OrigData");
    if(set_or_cil == 0) wks_sheet2 = WksIdxByName(pg, "Differentiation");
    else wks_sheet2 = WksIdxByName(pg, "Fit");
    wks_sheet3 = WksIdxByName(pg, "Parameters");
    wks_sheet4 = WksIdxByName(pg, "Filtration");
    num_col = ColIdxByUnit(wks_sheet1, (string)index_nach_otrezka).GetIndex();
	Dataset dsPila(wks_sheet1.Columns(0));
	
	is_str_numeric(get_str_from_str_list(wks_sheet2.Columns(num_col).GetUnits(), " ", 1), &otsech_nachala);
    is_str_numeric(get_str_from_str_list(wks_sheet2.Columns(num_col).GetUnits(), " ", 3), &otsech_konca);
	
	GETN_BOX( treeTest );

    if(wks_sheet4)
    {
    	is_str_numeric(get_str_from_str_list(wks_sheet4.Columns(num_col).GetUnits(), " ", 0), &Chast_to_Filtrate);
    	if(!Chast_to_Filtrate || Chast_to_Filtrate == NANUM) Chast_to_Filtrate = 0;
    	if(Chast_to_Filtrate !=0) Filtered_or_Not_1 = 1;
		else Chast_to_Filtrate = 0.01;
    	is_str_numeric(get_str_from_str_list(wks_sheet4.Columns(num_col).GetUnits(), " ", 2), &Chast_to_Filtrate_Diff);
    	is_str_numeric_integer(get_str_from_str_list(wks_sheet4.Columns(num_col).GetUnits(), " ", 4), &Poly_Order);
    	if(!Poly_Order || Poly_Order == NANUM || !get_str_from_str_list(wks_sheet4.Columns(num_col).GetUnits(), " ", 3).Match("PO")) 
    	{
    		Chast_to_Filtrate_Diff = Chast_to_Filtrate;
    		Poly_Order = 1;
    	}
    	else Filtered_or_Not = 1;
    }
    else
    {
    	Chast_to_Filtrate_Diff = 0;
    	Chast_to_Filtrate = 0.01;
    }
    
    GETN_SLIDEREDIT(decayT14, "Часть точек, которые будут отсечены от начала", otsech_nachala, "0.01|0.49|48")
    GETN_SLIDEREDIT(decayT15, "Часть точек, которые будут отсечены от конца", otsech_konca, "0.01|0.49|48")


    if(set_or_cil == 0) 
    {
    	GETN_SLIDEREDIT(decayT13, "Часть точек фильтрации", Chast_to_Filtrate, "0.01|0.99|98")
    	GETN_BEGIN_BRANCH(Details, "Сглаживание производной") GETN_CHECKBOX_BRANCH(Filtered_or_Not)
		if(Filtered_or_Not == 1) GETN_OPTION_BRANCH(GETNBRANCH_OPEN)
		GETN_COMBO(decayT1, "Степень сглаживающей функции", Poly_Order, "1|2|3|4|5")
		GETN_SLIDEREDIT(decayT2, "Часть точек фильтрации", Chast_to_Filtrate_Diff, "0.01|0.99|98")   
		GETN_END_BRANCH(Details)
    }
    else
    {
		GETN_BEGIN_BRANCH(Details1, "Фильтрация данных") GETN_CHECKBOX_BRANCH(Filtered_or_Not_1)
		if(Filtered_or_Not_1 == 1) GETN_OPTION_BRANCH(GETNBRANCH_OPEN)
		//GETN_COMBO(decayT1, "Степень сглаживающей функции", Poly_Order, "1|2|3|4|5")
		GETN_SLIDEREDIT(decayT2, "Часть точек фильтрации", Chast_to_Filtrate, "0.01|0.99|98") 
		GETN_END_BRANCH(Details1)
    }	
    
	if(0==GetNBox( treeTest, "Пересчет этого отрезка" )) return 0;
	
	TreeNode tn1;
	
    treeTest.GetNode("decayT14").GetValue(otsech_nachala);
	treeTest.GetNode("decayT15").GetValue(otsech_konca);
	otsech_nachala = (otsech_nachala == 0) ? 0.01 : otsech_nachala;
	otsech_konca = (otsech_konca == 0) ? 0.01 : otsech_konca;
	
	if(set_or_cil == 0) 
    {
		treeTest.GetNode("decayT13").GetValue(Chast_to_Filtrate);
		Chast_to_Filtrate = Chast_to_Filtrate/2;
		if(treeTest.Details.Use)
		{
			treeTest.GetNode("Details").GetNode("decayT1").GetValue(Poly_Order);
			treeTest.GetNode("Details").GetNode("decayT2").GetValue(Chast_to_Filtrate_Diff);
			Chast_to_Filtrate_Diff = Chast_to_Filtrate_Diff/2;
		}
		else Chast_to_Filtrate_Diff = 0;
    }
    else
    {
    	Chast_to_Filtrate_Diff = 0;
		if(treeTest.Details1.Use)
		{
			//treeTest.GetNode("Details1").GetNode("decayT1").GetValue(Poly_Order);
			treeTest.GetNode("Details1").GetNode("decayT2").GetValue(Chast_to_Filtrate);
			Chast_to_Filtrate = Chast_to_Filtrate/2;
		}
		else Chast_to_Filtrate = 0;
    }
	
    vector vParams(4), vec_X_stepennoi(wks_sheet1.Columns(0)), vec_Y_stepennoi(wks_sheet1.Columns(num_col)), vec_Y_stepennoi_posrednik;
    int new_size = size_otrezka*(1 - otsech_nachala - otsech_konca), ind_X = vec_X_stepennoi.GetSize(), ind_Y = vec_Y_stepennoi.GetSize();
    
    if(new_size < ind_Y)
    {
		vec_Y_stepennoi.SetSize(new_size);
		vec_X_stepennoi.SetSize(new_size);
		wks_sheet1.Columns(num_col).SetNumRows(new_size);
		wks_sheet2.Columns(num_col).SetNumRows(new_size);
    }
    else if(new_size > ind_Y)
    {
    	float shag = vec_X_stepennoi[ind_X-1] - vec_X_stepennoi[ind_X-2];
    	vec_Y_stepennoi.SetSize(new_size);
    	vec_X_stepennoi.SetSize(new_size);
    	for(int i = ind_X; i < vec_X_stepennoi.GetSize(); i++)
    		vec_X_stepennoi[i] = vec_X_stepennoi[i-1] + shag;
    	
    	Column colLol;
    	if(set_or_cil == 0)
    	{
			if(get_str_from_str_list(wks_graph_parent.GetPage().GetLongName(), " ", 1).Match("Т")) colLol = ColIdxByName(wks_OrigData, s_tst);
			else colLol = ColIdxByName(wks_OrigData, s_tssh);
    	}
    	else colLol = ColIdxByName(wks_OrigData, s_tc);
    	if(!colLol) 
    	{
    		MessageBox(GetWindow(), FuncErrors(222));
    		return 0;
    	}
    	vector vec_Tok_Orig(colLol);
    	
    	if(set_or_cil == 0) 
    		for(int j=0, k=index_nach_otrezka + size_otrezka*otsech_nachala; j<vec_Y_stepennoi.GetSize(); j++, k++) 
    			vec_Y_stepennoi[j] = vec_Tok_Orig[k]/resistance;
    	else 
    	{
    		for(int j=0, k=index_nach_otrezka + size_otrezka*otsech_nachala; j<vec_Y_stepennoi.GetSize(); j++, k++) 
    			vec_Y_stepennoi[j] = vec_Tok_Orig[k]/resistance;
    		turn_tok_upwards(vec_Y_stepennoi, 1, -1);
    	}
    }
    
    vec_Y_stepennoi_posrednik.SetSize(vec_Y_stepennoi.GetSize());
    vec_Y_stepennoi_posrednik = vec_Y_stepennoi;
    if(Chast_to_Filtrate != 0) 
    	ocmath_savitsky_golay(vec_Y_stepennoi, vec_Y_stepennoi_posrednik, vec_Y_stepennoi.GetSize(), vec_Y_stepennoi.GetSize()*Chast_to_Filtrate, -1, 5, 0, EDGEPAD_REFLECT);
	//vec_Y_stepennoi = vec_Y_stepennoi_posrednik;
	
	Dataset dsDiffed_Tok(wks_sheet2.Columns(num_col)), dsOrig_data(wks_sheet1.Columns(num_col)), dsFilted_data, dsTemp(wks_sheet3.Columns(2)), 
			dsMaxValue(wks_sheet3.Columns(1)), dsPeakVoltage(wks_sheet3.Columns(3)), dsEnergy(wks_sheet3.Columns(4)), dsX(wks_sheetLegend.Columns(0)),
			dsY(wks_sheetLegend.Columns(1)), dsSTOLBA(wks_sheet2.Columns(num_col));
	dsOrig_data = vec_Y_stepennoi;
	
	int nR1, nR2, up_down = 0;
	if(wks_sheet4 && Chast_to_Filtrate != 0)
	{
		dsFilted_data.Attach(wks_sheet4.Columns(num_col));
		dsFilted_data = vec_Y_stepennoi_posrednik;
	}
	else if(Chast_to_Filtrate != 0)
	{
		pg.AddLayer("Filtration");
		wks_sheet4 = pg.Layers(3);
		while( wks_sheet4.DeleteCol(0) );
		wks_sheet1.GetBounds(0, 1, nR1, -1, false);
		wks_sheet3.GetBounds(0, 1, nR2, -1, false);
		wks_sheet4.SetSize(nR1+1,nR2+2);
		wks_sheet4.Columns(0).SetLongName("pila");
		wks_sheet4.Columns(0).SetComments("0");
		wks_sheet4.Columns(0).SetType(OKDATAOBJ_DESIGNATION_X);
		wks_sheet4.Columns(0).SetUnits(wks_sheet1.Columns(0).GetUnits());
		Dataset ds_STOLBOV(wks_sheet4.Columns(0));
		ds_STOLBOV=vec_X_stepennoi;
		
		Tree trFormat;  
		trFormat.Root.CommonStyle.Fill.FillColor.nVal = 18; 
		DataRange dr;
		for(int i_2=1; i_2<nR2+2; i_2++)
		{
			//wks_sheet4.Columns(i_2).SetUnits(wks_sheet1.Columns(i_2).GetUnits());
			wks_sheet4.Columns(i_2).SetLongName(wks_sheet1.Columns(i_2).GetLongName());
			wks_sheet4.Columns(i_2).SetComments(wks_sheet1.Columns(i_2).GetComments());
			if(i_2 % 2 != 0)
			{
				dr.Add("X", wks_sheet4, 0, i_2, -1, i_2);
				if (0==dr.UpdateThemeIDs(trFormat.Root)) bool bRet = dr.ApplyFormat(trFormat, true, true);
				dr.Reset();
			}
		}
		dsFilted_data.Attach(wks_sheet4.Columns(num_col));
		dsFilted_data = vec_Y_stepennoi_posrednik;
		
		wks_sheet4.SetIndex(1);
	}
	else if(wks_sheet4)
	{
		dsFilted_data.Attach(wks_sheet4.Columns(num_col));
		dsFilted_data.SetSize(0);
	}
	
	if(Chast_to_Filtrate != 0 || Chast_to_Filtrate_Diff != 0)
	{
		string strFilt;
		strFilt.Format("%.2f", Chast_to_Filtrate*2);
		wks_sheet4.Columns(num_col).SetUnits(strFilt);
		if(Chast_to_Filtrate_Diff != 0) 
		{
			string str;
			str.Format("%.2f Pts %.2f PO %i", Chast_to_Filtrate*2, Chast_to_Filtrate_Diff*2, Poly_Order);
			wks_sheet4.Columns(num_col).SetUnits(str);
		}
	}
	else if(wks_sheet4) wks_sheet4.Columns(num_col).SetUnits((string)0);
	
	if(set_or_cil == 0)
    {
		if(Chast_to_Filtrate_Diff != 0)
		{
			ocmath_savitsky_golay(vec_Y_stepennoi_posrednik, vec_Y_stepennoi, vec_Y_stepennoi.GetSize(), vec_Y_stepennoi.GetSize()*Chast_to_Filtrate_Diff, -1, Poly_Order, 1, EDGEPAD_REFLECT);
			vec_Y_stepennoi = -vec_Y_stepennoi;
			dsDiffed_Tok = vec_Y_stepennoi;
		}
		else
		{
			ocmath_derivative( vec_X_stepennoi, vec_Y_stepennoi_posrednik, vec_Y_stepennoi.GetSize());
			vec_Y_stepennoi_posrednik = -vec_Y_stepennoi_posrednik;
			dsDiffed_Tok = vec_Y_stepennoi_posrednik;
		}
    }
    else
    {
    	NLFitSession FitSession;
    	FitSession.SetFunction("Gauss");
		FitSession.SetMaxNumIter(10);
		
		FitSession.SetData(vec_Y_stepennoi_posrednik, vec_X_stepennoi);
			
		FitSession.ParamsInitValues();
		
		int     nOutcome;
		FitSession.Fit(&nOutcome, false, false);
		
		vector vParamValues, vErrors;
		FitSession.GetFitResultsParams(vParamValues, vErrors);
		
		FitSession.GetYFromX(vec_X_stepennoi, vec_Y_stepennoi_posrednik, vec_Y_stepennoi_posrednik.GetSize());
		
		dsDiffed_Tok = vec_Y_stepennoi_posrednik;
		
		dsTemp[num_col-1] = vParamValues[2];
		
		up_down = (vParamValues[3] >= 0) ? 0 : 1;
    }
    
    vector v_whm_w_peak(10);
	/*
	вектор v_whm_w_peak:
		whm
		-[0] ширина на полувысоте
		-[1] Y на полувысоте
		-[2] X на полувысоте слева
		-[3] X на полувысоте справа
		W
		-[4] значение коэффициента W
		-[5] Y при W
		-[6] X при W слева
		-[7] X при W справа
		Peak
		-[8] Y пика
		-[9] X пика
	*/
	
	v_whm_w_peak = find_whm_Wparam_LRmidindxs(vec_X_stepennoi, dsDiffed_Tok, set_or_cil, dsTemp[num_col-1], up_down);
		
	if(set_or_cil == 0)
	{
		dsMaxValue[num_col-1] = v_whm_w_peak[8];
		dsPeakVoltage[num_col-1] = v_whm_w_peak[9];
		dsTemp[num_col-1] = v_whm_w_peak[0];
	}
	else
	{
		dsMaxValue[num_col-1] = v_whm_w_peak[8];
		dsPeakVoltage[num_col-1] = v_whm_w_peak[9];
		dsEnergy[num_col-1] = v_whm_w_peak[0];
	}
	
	dsX[0] = vec_X_stepennoi[0];
    dsX[1] = vec_X_stepennoi[dsSTOLBA.GetSize()-1];
	dsY[0] = dsY[1] = dsMaxValue[num_col-1];
    
    if(set_or_cil == 0)
	{
		dsY.Attach(wks_sheetLegend.Columns(2));
		dsY[0] = dsY[1] = v_whm_w_peak[1];
	}
	else
	{
		dsY.Attach(wks_sheetLegend.Columns(2));
		dsY[0] = dsY[1] = v_whm_w_peak[5];
	}
    
	dsX.Attach(wks_sheetLegend.Columns(3));
    dsY.Attach(wks_sheetLegend.Columns(4));
    dsSTOLBA.Attach(wks_sheet2.Columns(num_col));
    dsY[0] = min(dsSTOLBA);
    dsY[1] = max(dsSTOLBA);

	if(set_or_cil == 0)
		dsX[0] = dsX[1] = v_whm_w_peak[2];
	else 
		dsX[0] = dsX[1] = v_whm_w_peak[6];
	
    dsX.Attach(wks_sheetLegend.Columns(5));
    dsY.Attach(wks_sheetLegend.Columns(6));
    dsY[0] = min(dsSTOLBA);
    dsY[1] = max(dsSTOLBA);

    if(set_or_cil == 0)
		dsX[0] = dsX[1] = v_whm_w_peak[3];
	else 
		dsX[0] = dsX[1] = v_whm_w_peak[7];
	
	string otsecheniya;
	otsecheniya.Format("S %.2f E %.2f", otsech_nachala, otsech_konca);
	wks_sheet2.Columns(num_col).SetUnits(otsecheniya);
	
	if(ind_X<vec_X_stepennoi.GetSize())
	{
		dsPila = vec_X_stepennoi;
		dsPila.Attach(wks_sheet2.Columns(0));
		dsPila = vec_X_stepennoi;
		if(wks_sheet4)
		{
			dsPila.Attach(wks_sheet4.Columns(0));
			dsPila = vec_X_stepennoi;
		}
	}
	
    wks_sheet1.GetBounds(nR1, 1, nR2, -1, false);
    wks_sheet1.SetSize(nR2+1, -1);
    wks_sheet2.SetSize(nR2+1, -1);
    if(wks_sheet4) wks_sheet4.SetSize(nR2+1, -1);
	
	return -1;
}

void pererisovka_grafika_setoch(GraphLayer graph, int num_col, Worksheet wks_graph_parent, Worksheet wks_Legend)
{
	Worksheet wks_sheet1, wks_sheet2, wks_sheetLegend, wks_sheet4;
	Page pg =  wks_graph_parent.GetPage(), pgLegend =  wks_Legend.GetPage();
	int set_or_cil;
	if(pg.GetLongName().Match("Сетка*")) set_or_cil = 0;
    if(pg.GetLongName().Match("Цилиндр*")||pg.GetLongName().Match("Магнит*")) set_or_cil = 1;
	if(set_or_cil == 0) wks_sheet1 = WksIdxByName(pg, "Differentiation");
    else wks_sheet1 = WksIdxByName(pg, "Fit");
    wks_sheet2 = WksIdxByName(pg, "Parameters");
    wks_sheet4 = WksIdxByName(pg, "Filtration");
    wks_sheetLegend = pgLegend.Layers(1);
    
    Dataset dsTemp(wks_sheet2.Columns(2)), dsMaxVals(wks_sheet2.Columns(1)), dsX(wks_sheetLegend.Columns(0)), dsY(wks_sheetLegend.Columns(1)), dsSTOLBA(wks_sheet1.Columns(num_col)), dsPILY(wks_sheet1.Columns(0));
    double He_or_Ar;
    
    int nPlots = graph.DataPlots.Count();
    if(wks_sheet4 && set_or_cil == 1)
	{
		if(wks_sheet4.Columns(num_col).GetUpperBound() > 0)
		{
			if(nPlots < 8)
			{
				DataRange dr5;
				dr5.Add(wks_sheet4, 0, "X");
				dr5.Add(wks_sheet4, num_col, "Y");
				graph.AddPlot(dr5, IDM_PLOT_LINE);
				DataPlot dp5 = graph.DataPlots(7);
				Tree tr4;
				tr4 = dp5.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
				tr4.Root.Line.Width.dVal=2.5;
				tr4.Root.Line.Color.nVal=2;
				dp5.ApplyFormat(tr4, true, true);
				vector<uint> vecUI = {0,7,1,2,3,4,5,6};
				graph.ReorderPlots(vecUI);
			}
		}
		else if(nPlots == 8) graph.RemovePlot(1);
	}

    string str;
    wks_Legend.GetCell(2, 1, str);
    if(str.Match("He"))He_or_Ar=0;
    else He_or_Ar=1;

	GraphLayer gl1 = graph.GetPage().Layers(0), gl2 = graph.GetPage().Layers(1);
	gl1.Rescale();
	if(set_or_cil == 0) gl2.Rescale();
	GraphPage gp = gl1.GetPage();
	legend_combine(gp);
	
	Axis axisX = gl1.XAxis;
	if(dsPILY[0] >= 0) axisX.Scale.From.dVal = 0;
	else axisX.Scale.From.dVal = dsPILY[0];
	axisX.Scale.To.dVal = dsPILY[dsPILY.GetSize()-1];
	
	wks_Legend.SetCell(0, 1, dsTemp[num_col-1]);
	dsTemp.Attach( wks_sheet2.Columns(3));
	wks_Legend.SetCell(1, 1, dsTemp[num_col-1]);
	wks_Legend.AutoSize();
	
	if(set_or_cil == 0) create_labels(1);
    else create_labels(2);
}

int Get_OriginData_wks(Worksheet &wks_graph_parent, Worksheet &wks_OrigData, Worksheet &wks_Legend)
{
	GraphLayer graphActive = Project.ActiveLayer(), graph = graphActive.GetPage().Layers(0);
	if(!graph) 
	{
		MessageBox(GetWindow(), "!graph");
		return 0;
	}
	vector<uint> vUIDs;
    if(graph.GetConnectedObjects(vUIDs)<=1) 
    {
    	MessageBox(GetWindow(), "Не найден родительский объект графика или книга легенды");
		return 0;
    }
    wks_graph_parent = Project.GetObject(vUIDs[0]);
    if(!wks_graph_parent) 
    {
    	MessageBox(GetWindow(), "Не найден родительский объект графика");
		return 0;
    }
    wks_Legend = Project.GetObject(vUIDs[1]);
    if(!wks_Legend) 
    {
    	MessageBox(GetWindow(), "Не найдена книга легенды");
		return 0;
    }
    vUIDs.RemoveAll();
    if(wks_graph_parent.GetConnectedObjects(vUIDs)<=0)
    {
    	MessageBox(GetWindow(), "Не найдена книга с оригинальными данными");
		return 0;
    }
    for(int ii=0; ii<vUIDs.GetSize(); ii++)
    {
        wks_OrigData = Project.GetObject(vUIDs[ii]);
        if(wks_OrigData) return -1;
    }
	if(!wks_OrigData)
	{
		MessageBox(GetWindow(), "Недействительна книга с оригинальными данными");
		return 0;
	}
    return -1;
}

#define TOL 1e-30 /* smallest value allowed in cholesky_decomp() */

void levmarq(vector vec_Y_stepennoi, vector vec_X_stepennoi, vector &vParams, int Num_iter)
{
	double A=vParams[0],B=vParams[1],C=vParams[2],D=vParams[3];
	int x,i,j,it,ill, npar = 2;
	double lambda = 0.001,up = 10,down = 1/up,mult,err = 0,newerr = 0,derr = 0,target_derr = 10^-12, new_C, new_D;
	matrix Hessian,ch;
	vector grad,drvtv,delta;
	Hessian.SetSize(2, 2);
	ch.SetSize(2, 2);
	grad.SetSize(2);
	drvtv.SetSize(2);
	delta.SetSize(2);

	/* calculate the initial error ("chi-squared") */
	err = error_func(vec_Y_stepennoi, vec_X_stepennoi, A, B, C, D);

	/* main iteration */
	for (it = 0; it < Num_iter; it++) 
	{
		/* calculate the approximation to the Hessian and the "derivative" drvtv */
		drvtv[0] = 0;
		drvtv[1] = 0;	
		Hessian[0][0] = 0;
		Hessian[0][1] = 0;
		Hessian[1][0] = 0;
		Hessian[1][1] = 0;
		
		for (x = vec_Y_stepennoi.GetSize()*0.5; x < vec_Y_stepennoi.GetSize(); x++) 
		{
			grad[0] = dfxdC(vec_X_stepennoi[x], D);
			grad[1] = dfxdD(vec_X_stepennoi[x], C, D);
			
			drvtv[0] += (vec_Y_stepennoi[x] - fx(vec_X_stepennoi[x], A, B, C, D))*grad[0];
			drvtv[1] += (vec_Y_stepennoi[x] - fx(vec_X_stepennoi[x], A, B, C, D))*grad[1];
			
			Hessian[0][0] += grad[0]*grad[0];
			Hessian[0][1] += grad[0]*grad[1];
			Hessian[1][0] += grad[1]*grad[0];
			Hessian[1][1] += grad[1]*grad[1];
		}

		/*  make a step "delta."  If the step is rejected, increase lambda and try again */
		mult = 1 + lambda;
		ill = 1; /* ill-conditioned? */
		while (ill && (it < Num_iter)) 
		{
			for (i = 0; i < npar; i++) Hessian[i][i] = Hessian[i][i]*mult;

			ill = cholesky_decomp(ch, Hessian);

			if (!ill) 
			{
				solve_axb_cholesky(ch, delta, drvtv);
				new_C = C + delta[0];
				new_D = D + delta[1];
				newerr = error_func(vec_Y_stepennoi, vec_X_stepennoi, A, B, new_C, new_D);
				derr = newerr - err;
				ill = (derr > 0);
			} 

			if (ill) 
			{
				mult = (1 + lambda*up)/(1 + lambda);
				lambda *= up;
				it++;
			}
		}
		C = new_C;
		D = new_D;
		err = newerr;
		lambda *= down;  

		if ((!ill) && (-derr < target_derr)) break;
	}
	vParams[2] = C;
	vParams[3] = D;
}

/* calculate the error function (chi-squared) */
double error_func(vector vec_Y_stepennoi, vector vec_X_stepennoi, double A, double B, double C, double D)
{
	double res,e=0;
	for (int i = vec_Y_stepennoi.GetSize()*0.5; i < vec_Y_stepennoi.GetSize(); i++) 
	{
		res = fx(vec_X_stepennoi[i], A, B, C, D) - vec_Y_stepennoi[i];
		e += res*res;
	}
	return e;
}

/* solve the equation Ax=b for a symmetric positive-definite matrix A,
   using the Cholesky decomposition A=LL^T.  The matrix L is passed in "ch".
   Elements above the diagonal are ignored. */
void solve_axb_cholesky(matrix ch, vector &delta, vector drvtv)
{
	int i,j, npar = 2;
	double sum;
	
	/* solve L*y = b for y (where delta[] is used to store y) */

	for (i = 0; i < npar; i++) 
	{
		sum = 0;
		for (j = 0; j < i; j++) sum += ch[i][j] * delta[j];
		delta[i] = (drvtv[i] - sum)/ch[i][i];      
	}

	/* solve L^T*delta = y for delta (where delta[] is used to store both y and delta) */

	for (i = npar-1; i >= 0; i--) 
	{
		sum = 0;
		for (j = i+1; j < npar; j++) sum += ch[j][i] * delta[j];
		delta[i] = (delta[i] - sum)/ch[i][i];      
	}
}

/* This function takes a symmetric, positive-definite matrix "Hessian" and returns
   its (lower-triangular) Cholesky factor in "ch".  Elements above the
   diagonal are neither used nor modified.  The same array may be passed
   as both ch and Hessian, in which case the decomposition is performed in place. */
bool cholesky_decomp(matrix &ch, matrix Hessian)
{
	int i,j,k, npar = 2;
	double sum;

	for (i = 0; i < npar; i++) 
	{
		for (j = 0; j < i; j++) 
		{
			sum = 0;
			for (k = 0; k < j; k++) sum += ch[i][k] * ch[j][k];
			ch[i][j] = (Hessian[i][j] - sum)/ch[j][j];
		}
		sum = 0;
		for (k = 0; k < i; k++) sum += ch[i][k] * ch[i][k];
		sum = Hessian[i][i] - sum;
		if (sum < TOL) return -1; /* not positive-definite */
		ch[i][i] = sqrt(sum);
	}
	return 0;
}

void cilinder()
{
	Worksheet wks = Project.ActiveLayer(), wks_slave;
    if(!wks)
    {
    	MessageBox(GetWindow(), "wks!");
    	return;
    }
    
	if(wks.GetPage().GetLongName().Match("Зонд*") || wks.GetPage().GetLongName().Match("Сетка*") 
		|| wks.GetPage().GetLongName().Match("Цилиндр*") || wks.GetPage().GetLongName().Match("Магнит*"))
	{
		MessageBox(GetWindow(), "Не тот лист! Выберите лист с результатами эксперимента");
		return;
	}

    vector<string> vBooksNames, vErrsBooksNames;
	foreach(WorksheetPage pg in Project.WorksheetPages)
	{
		int BeginVal;
		is_str_numeric_integer(get_str_from_str_list(pg.GetLongName(), "_", 0), &BeginVal);
		if(BeginVal>10000) vBooksNames.Add(pg.GetLongName());
	}
	
    GETN_BOX( treeTest );
    //GETN_NUM(decayT2, "Частота пилы, Гц", ch_pil)
    
    GETN_BEGIN_BRANCH(Start, "Уравнение преобразования напряжения")
    GETN_MULTI_COLS_BRANCH(6, DYNA_MULTI_COLS_COMAPCT) GETN_OPTION_BRANCH(GETNBRANCH_OPEN)
		
		GETN_CHECK(decayT3,"Обработка данных c цилиндрического", true)
		GETN_NUM(decayT5,"U =", c_A) GETN_EDITOR_SIZE_USER_OPTION("#5")
		GETN_NUM(decayT6, "*", c_B) GETN_EDITOR_SIZE_USER_OPTION("#5")
		GETN_NUM(decayT7,"*", c_C) GETN_EDITOR_SIZE_USER_OPTION("#5")
		GETN_NUM(decayT8, "/", c_D) GETN_EDITOR_SIZE_USER_OPTION("#5")
		GETN_NUM(decayT9, "-", c_E) GETN_EDITOR_SIZE_USER_OPTION("#5")
		GETN_CHECK(decayT10,"Обработка данных с магнитного", false)
		GETN_NUM(decayT12,"U =", m_A) GETN_EDITOR_SIZE_USER_OPTION("#5")
		GETN_NUM(decayT13, "*", m_B) GETN_EDITOR_SIZE_USER_OPTION("#5")
		GETN_NUM(decayT14,"*", m_C) GETN_EDITOR_SIZE_USER_OPTION("#5")
		GETN_NUM(decayT15, "/", m_D) GETN_EDITOR_SIZE_USER_OPTION("#5")
		GETN_NUM(decayT16, "-", m_E) GETN_EDITOR_SIZE_USER_OPTION("#5")
		
    GETN_END_BRANCH(Start)
    
    GETN_BEGIN_BRANCH(Details2, "Задать частоты дискретизации(по умолчанию автоопределение)") GETN_CHECKBOX_BRANCH(0)
    //GETN_OPTION_BRANCH(GETNBRANCH_CLOSED)
	GETN_NUM(decayT8, "Частота дискретизации пилы, Гц", ch_dsc_p)
    GETN_NUM(decayT9, "Частота дискретизации цилиндра, Гц", ch_dsc_t)
    GETN_END_BRANCH(Details2)
    
    GETN_BEGIN_BRANCH(Details3, "Задать начало и конец обработки(по умолчанию автоопределение)") GETN_CHECKBOX_BRANCH(0)
    //GETN_OPTION_BRANCH(GETNBRANCH_CLOSED)
    GETN_SLIDEREDIT(decayT1, "Время начала, с", 0.3, "0.0|5.0|50") 
    GETN_SLIDEREDIT(decayT2, "Время конца, с", 1.5, "0.0|5.0|50") 
    GETN_END_BRANCH(Details3)
    
    GETN_SLIDEREDIT(decayT14, "Часть точек, которые будут отсечены от начала", 0.1, "0.1|0.49|39")
    GETN_SLIDEREDIT(decayT15, "Часть точек, которые будут отсечены от конца", 0.1, "0.1|0.49|39")
    
    //GETN_NUM(decayT1, "Сопротивление на цилиндре, Ом", r_cil)
    GETN_RADIO_INDEX_EX(decayT11, "Тип импульса", 2, "СВЧ|ВЧ|ИЦР")
    GETN_RADIO_INDEX_EX(decayT13, "Рабочее тело", 0, "He|Ar")
    
    GETN_BEGIN_BRANCH(Details1, "Фильтрация данных") GETN_CHECKBOX_BRANCH(0)
    GETN_OPTION_BRANCH(GETNBRANCH_OPEN)
    GETN_SLIDEREDIT(decayT14, "Часть точек фильтрации", 0.2, "0.01|0.99|98") 
    GETN_END_BRANCH(Details1)
    
    GETN_BEGIN_BRANCH(Details, "Нужно обработать несколько книг") GETN_CHECKBOX_BRANCH(0)
    for(int i_1 = 0; i_1<vBooksNames.GetSize(); i_1++) 
    {
		GETN_CHECK(NodeName, vBooksNames[i_1], true)
    }
    GETN_END_BRANCH(Details)
    
    if(0==GetNBox( treeTest, "Обработка данных с цилиндра/магнита" )) return;
    
    double resistance = r_cil,chastota_pily = ch_pil,chastota_discret_pily,chastota_discret_CILIND, Chast_to_Filtrate;
    int Obr_C, Obr_M, SVCH_or_VCH, He_or_Ar, True_or_False, Error = -1, UseSameName_or_Not_PILA_C, UseSameName_or_Not_PILA_M, UseSameName_or_Not_CILIND, 
		UseSameName_or_Not_MAGNIT, z_s_c = 2;
    vector vecErrs;
    TreeNode tn1;
    
    //////////////////////////// СЧИТЫВАЕМ ДАННЫЕ ИЗ ДИАЛОГА ////////////////////////////
    treeTest.GetNode("Start").GetNode("decayT3").GetValue(Obr_C);
    treeTest.GetNode("Start").GetNode("decayT5").GetValue(c_A);
	treeTest.GetNode("Start").GetNode("decayT6").GetValue(c_B);
	treeTest.GetNode("Start").GetNode("decayT7").GetValue(c_C);
	treeTest.GetNode("Start").GetNode("decayT8").GetValue(c_D);
	treeTest.GetNode("Start").GetNode("decayT9").GetValue(c_E);
	treeTest.GetNode("Start").GetNode("decayT10").GetValue(Obr_M);
    treeTest.GetNode("Start").GetNode("decayT12").GetValue(m_A);
	treeTest.GetNode("Start").GetNode("decayT13").GetValue(m_B);
	treeTest.GetNode("Start").GetNode("decayT14").GetValue(m_C);
	treeTest.GetNode("Start").GetNode("decayT15").GetValue(m_D);
	treeTest.GetNode("Start").GetNode("decayT16").GetValue(m_E);
	//treeTest.GetNode("decayT1").GetValue(resistance);
	//treeTest.GetNode("decayT2").GetValue(chastota_pily);
	treeTest.GetNode("decayT8").GetValue(chastota_discret_pily);
	treeTest.GetNode("decayT9").GetValue(chastota_discret_CILIND);
	treeTest.GetNode("decayT11").GetValue(SVCH_or_VCH);
	treeTest.GetNode("decayT13").GetValue(He_or_Ar);
	
	treeTest.GetNode("decayT14").GetValue(otsech_nachala);
	treeTest.GetNode("decayT15").GetValue(otsech_konca);
	otsech_nachala = (otsech_nachala == 0) ? 0.01 : otsech_nachala;
	otsech_konca = (otsech_konca == 0) ? 0.01 : otsech_konca;
	
	if(treeTest.Details1.Use)
	{
		treeTest.GetNode("Details1").GetNode("decayT14").GetValue(Chast_to_Filtrate);
		Chast_to_Filtrate = Chast_to_Filtrate/2;
	}
	else Chast_to_Filtrate = 0;
	if(treeTest.Details3.Use)
	{
		treeTest.GetNode("Details3").GetNode("decayT1").GetValue(st_time_end_time[0]);
		treeTest.GetNode("Details3").GetNode("decayT2").GetValue(st_time_end_time[1]);
		if(st_time_end_time[0] >= st_time_end_time[1])
		{
			MessageBox(GetWindow(), "Ошибка. Исправьте время начала и конца обработки. Вы выбрали от " + st_time_end_time[0] + " до " + st_time_end_time[1]);
			return;
		}
	}
	//////////////////////////// СЧИТЫВАЕМ ДАННЫЕ ИЗ ДИАЛОГА ////////////////////////////
	
	string ReportString_CILIND, ReportString_PILA_C, ReportString_PILA_M, ReportString_MAGNIT;
   	Column colPILA_C, colPILA_M, colTOK_CILIND, colTOK_MAGNIT;
   	if(Obr_C == 1)
   	{
   		colPILA_C = ColIdxByName(wks, s_pc);
		if(!colPILA_C)
		{
			GETN_BOX( treeTest1 );
			GETN_STR(STR, "!Не найден столбец с именем " + "\"" + s_pc + "\"" + "!", "") GETN_INFO
			GETN_INTERACTIVE(interactive, "Выберите столбец с нужными данными", "[Book1]Sheet1!B")
			GETN_OPTION_INTERACTIVE_CONTROL(ICOPT_RESTRICT_TO_ONE_DATA) 
			if(treeTest.Details.Use)
			{
				GETN_RADIO_INDEX_EX(decayT1, "Использовать колонку с таким же именем для расчета остальных книг", 0, "Да|Нет")
			}
			if(0==GetNBox( treeTest1, "Ошибка!" )) return;
			
			treeTest1.GetNode("interactive").GetValue(ReportString_PILA_C);
			treeTest1.GetNode("decayT1").GetValue(UseSameName_or_Not_PILA_C);
			ReportString_PILA_C = get_str_from_str_list(ReportString_PILA_C, "\"", 1);
			colPILA_C = ColIdxByName(wks, ReportString_PILA_C);
			if(!colPILA_C)
			{
				MessageBox(GetWindow(), "Ошибка. Возможно, выбранная вами колонка пустая");
				return;
			}
			colPILA_C.SetLongName(s_pc);
			UseSameName_or_Not_PILA_C = UseSameName_or_Not_PILA_C + 2;
		}
		colTOK_CILIND = ColIdxByName(wks, s_tc);
		if(!colTOK_CILIND)
		{	
			GETN_BOX( treeTest1 );
			GETN_STR(STR, "!Не найден столбец с именем " + "\"" + s_tc + "\"" + "!", "") GETN_INFO
			GETN_INTERACTIVE(interactive, "Выберите столбец с нужными данными", "[Book1]Sheet1!B")
			GETN_OPTION_INTERACTIVE_CONTROL(ICOPT_RESTRICT_TO_ONE_DATA) 
			if(treeTest.Details.Use) 
			{
				GETN_RADIO_INDEX_EX(decayT1, "Использовать колонку с таким же именем для расчета остальных книг", 0, "Да|Нет")
			}
			if(0==GetNBox( treeTest1, "Ошибка!" )) return;
			
			treeTest1.GetNode("interactive").GetValue(ReportString_CILIND);
			treeTest1.GetNode("decayT1").GetValue(UseSameName_or_Not_CILIND);
			ReportString_CILIND = get_str_from_str_list(ReportString_CILIND, "\"", 1);
			colTOK_CILIND = ColIdxByName(wks, ReportString_CILIND);
			if(!colTOK_CILIND)
			{
				MessageBox(GetWindow(), "Ошибка. Возможно, выбранная вами колонка пустая");
				return;
			}
			colTOK_CILIND.SetLongName(s_tc);
			UseSameName_or_Not_CILIND = UseSameName_or_Not_CILIND + 2;
		}
   	}
   	if(Obr_M == 1)
   	{
   		colPILA_M = ColIdxByName(wks, s_pm);
		if(!colPILA_M)
		{
			GETN_BOX( treeTest1 );
			GETN_STR(STR, "!Не найден столбец с именем " + "\"" + s_pm + "\"" + "!", "") GETN_INFO
			GETN_INTERACTIVE(interactive, "Выберите столбец с нужными данными", "[Book1]Sheet1!B")
			GETN_OPTION_INTERACTIVE_CONTROL(ICOPT_RESTRICT_TO_ONE_DATA) 
			if(treeTest.Details.Use)
			{
				GETN_RADIO_INDEX_EX(decayT1, "Использовать колонку с таким же именем для расчета остальных книг", 0, "Да|Нет")
			}
			if(0==GetNBox( treeTest1, "Ошибка!" )) return;
			
			treeTest1.GetNode("interactive").GetValue(ReportString_PILA_M);
			treeTest1.GetNode("decayT1").GetValue(UseSameName_or_Not_PILA_M);
			ReportString_PILA_M = get_str_from_str_list(ReportString_PILA_M, "\"", 1);
			colPILA_M = ColIdxByName(wks, ReportString_PILA_M);
			if(!colPILA_M)
			{
				MessageBox(GetWindow(), "Ошибка. Возможно, выбранная вами колонка пустая");
				return;
			}
			colPILA_M.SetLongName(s_pm);
			UseSameName_or_Not_PILA_M = UseSameName_or_Not_PILA_M + 2;
		}
		colTOK_MAGNIT = ColIdxByName(wks, s_tm);
		if(!colTOK_MAGNIT)
		{	
			GETN_BOX( treeTest1 );
			GETN_STR(STR, "!Не найден столбец с именем " + "\"" + s_tm + "\"" + "!", "") GETN_INFO
			GETN_INTERACTIVE(interactive, "Выберите столбец с нужными данными", "[Book1]Sheet1!B")
			GETN_OPTION_INTERACTIVE_CONTROL(ICOPT_RESTRICT_TO_ONE_DATA) 
			if(treeTest.Details.Use) 
			{
				GETN_RADIO_INDEX_EX(decayT1, "Использовать колонку с таким же именем для расчета остальных книг", 0, "Да|Нет")
			}
			if(0==GetNBox( treeTest1, "Ошибка!" )) return;
			
			treeTest1.GetNode("interactive").GetValue(ReportString_MAGNIT);
			treeTest1.GetNode("decayT1").GetValue(UseSameName_or_Not_MAGNIT);
			ReportString_MAGNIT = get_str_from_str_list(ReportString_MAGNIT, "\"", 1);
			colTOK_MAGNIT = ColIdxByName(wks, ReportString_MAGNIT);
			if(!colTOK_MAGNIT)
			{
				MessageBox(GetWindow(), "Ошибка. Возможно, выбранная вами колонка пустая");
				return;
			}
			colTOK_MAGNIT.SetLongName(s_tm);
			UseSameName_or_Not_MAGNIT = UseSameName_or_Not_MAGNIT + 2;
		}
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
				Column colTOK_CILIND_slave, colPILA_C_slave, colTOK_MAGNIT_slave, colPILA_M_slave;
				Page pg1 = Project.FindPage(vBooksNames[i_2]);
				if(pg1.Layers(0) != wks) 
				{
					wks_slave = pg1.Layers(0);
					if(UseSameName_or_Not_CILIND == 2) 
					{
						colTOK_CILIND_slave = ColIdxByName(wks_slave, ReportString_CILIND);
						if(!colTOK_CILIND_slave)
						{
							vecErrs.Add(222);
							vErrsBooksNames.Add("Обработка цилинда "+vBooksNames[i_2]+": ");
							continue;
						}
					}
					if(UseSameName_or_Not_PILA_C == 2) 
					{
						colPILA_C_slave = ColIdxByName(wks_slave, ReportString_PILA_C);
						if(!colPILA_C_slave)
						{
							vecErrs.Add(333);
							vErrsBooksNames.Add("Пила цилиндра "+vBooksNames[i_2]+": ");
							continue;
						}
					}
					if(UseSameName_or_Not_MAGNIT == 2) 
					{
						colTOK_CILIND_slave = ColIdxByName(wks_slave, ReportString_MAGNIT);
						if(!colTOK_MAGNIT_slave)
						{
							vecErrs.Add(222);
							vErrsBooksNames.Add("Обработка магнитного "+vBooksNames[i_2]+": ");
							continue;
						}
					}
					if(UseSameName_or_Not_PILA_M == 2) 
					{
						colPILA_M_slave = ColIdxByName(wks_slave, ReportString_PILA_M);
						if(!colPILA_M_slave)
						{
							vecErrs.Add(333);
							vErrsBooksNames.Add("Пила магнитного "+vBooksNames[i_2]+": ");
							continue;
						}
					}
				}
				else
				{
					wks_slave = wks;
					colTOK_CILIND_slave = colTOK_CILIND;
					colPILA_C_slave = colPILA_C;
					colTOK_MAGNIT_slave = colTOK_MAGNIT;
					colPILA_M_slave = colPILA_M;
				}
				colTOK_CILIND_slave.SetLongName(s_tc);
				colTOK_MAGNIT_slave.SetLongName(s_tm);
				colPILA_C_slave.SetLongName(s_pc);
				colPILA_M_slave.SetLongName(s_pm);
				
				if(Obr_C == 1 && Error < 0) 
				{
					if(!treeTest.Details2.Use)
					{
						Error = auto_chdsc(chastota_discret_pily, chastota_discret_CILIND, colPILA_C_slave, colTOK_CILIND_slave, ch_pil, z_s_c, wks_slave, 0);
						if( Error != -1 ) 
						{
							vecErrs.Add(Error);
							if(Error != 333) vErrsBooksNames.Add("Обработка магнитного "+vBooksNames[i_2]+": ");
							else vErrsBooksNames.Add("Пила магнитного "+vBooksNames[i_2]+": ");
						}
						Error = -1;
					}
					else
					{
						treeTest.GetNode("Details2").GetNode("decayT8").GetValue(chastota_discret_pily);
						treeTest.GetNode("Details2").GetNode("decayT9").GetValue(chastota_discret_CILIND);
					}
					
					Error = zond_obrab(resistance, chastota_pily, chastota_discret_pily, chastota_discret_CILIND, SVCH_or_VCH, He_or_Ar, 
											wks_slave, colPILA_C_slave, colTOK_CILIND_slave, Chast_to_Filtrate, z_s_c, 0, S1, c_A, c_B, c_C, c_D, c_E);
					if(Error == 404) return;
					if(Error != -1) 
					{
						vecErrs.Add(Error);
						if(Error != 333) vErrsBooksNames.Add("Обработка цилиндра "+vBooksNames[i_2]+": ");
						else vErrsBooksNames.Add("Пила цилиндра "+vBooksNames[i_2]+": ");
					}
					Error = -1;
				}
				if(Obr_M == 1 && Error < 0) 
				{
					if(!treeTest.Details2.Use)
					{
						Error = auto_chdsc(chastota_discret_pily, chastota_discret_CILIND, colPILA_M_slave, colTOK_MAGNIT_slave, ch_pil, z_s_c, wks_slave, 1);
						if( Error != -1 ) 
						{
							vecErrs.Add(Error);
							if(Error != 333) vErrsBooksNames.Add("Обработка магнитного "+vBooksNames[i_2]+": ");
							else vErrsBooksNames.Add("Пила магнитного "+vBooksNames[i_2]+": ");
						}
						Error = -1;
					}
					else
					{
						treeTest.GetNode("Details2").GetNode("decayT8").GetValue(chastota_discret_pily);
						treeTest.GetNode("Details2").GetNode("decayT9").GetValue(chastota_discret_CILIND);
					}
					
					Error = zond_obrab(resistance, chastota_pily, chastota_discret_pily, chastota_discret_CILIND, SVCH_or_VCH, He_or_Ar, 
											wks_slave, colPILA_M_slave, colTOK_MAGNIT_slave, Chast_to_Filtrate, z_s_c, 1, S2, m_A, m_B, m_C, m_D, m_E);
					if(Error == 404) return;
					if(Error != -1) 
					{
						vecErrs.Add(Error);
						if(Error != 333) vErrsBooksNames.Add("Обработка магнитного "+vBooksNames[i_2]+": ");
						else vErrsBooksNames.Add("Пила магнитного "+vBooksNames[i_2]+": ");
					}
					Error = -1;
				}
			}
		}
	}
	else 
	{
		if(Obr_C == 1 && Error < 0) 
		{
			if(!treeTest.Details2.Use)
			{
				Error = auto_chdsc(chastota_discret_pily, chastota_discret_CILIND, colPILA_C, colTOK_CILIND, ch_pil, z_s_c, wks, 0);
				if( Error != -1 ) 
				{
					MessageBox(GetWindow(), FuncErrors(Error));
					return;
				}
				Error = -1;
			}
			else
			{
				treeTest.GetNode("Details2").GetNode("decayT8").GetValue(chastota_discret_pily);
				treeTest.GetNode("Details2").GetNode("decayT9").GetValue(chastota_discret_CILIND);
			}
			
			Error = zond_obrab(resistance, chastota_pily, chastota_discret_pily, chastota_discret_CILIND, SVCH_or_VCH, He_or_Ar, 
								wks, colPILA_C, colTOK_CILIND, Chast_to_Filtrate, z_s_c, 0, S1, c_A, c_B, c_C, c_D, c_E);
			if(Error == 404) return;
			if( Error != -1 ) 
			{
				MessageBox(GetWindow(), FuncErrors(Error));
				return;
			}
			Error = -1;
		}
		if(Obr_M == 1 && Error < 0) 
		{
			if(!treeTest.Details2.Use)
			{
				Error = auto_chdsc(chastota_discret_pily, chastota_discret_CILIND, colPILA_M, colTOK_MAGNIT, ch_pil, z_s_c, wks, 1);
				if( Error != -1 ) 
				{
					MessageBox(GetWindow(), FuncErrors(Error));
					return;
				}
				Error = -1;
			}
			else
			{
				treeTest.GetNode("Details2").GetNode("decayT8").GetValue(chastota_discret_pily);
				treeTest.GetNode("Details2").GetNode("decayT9").GetValue(chastota_discret_CILIND);
			}
			
			Error = zond_obrab(resistance, chastota_pily, chastota_discret_pily, chastota_discret_CILIND, SVCH_or_VCH, He_or_Ar, 
								wks, colPILA_M, colTOK_MAGNIT, Chast_to_Filtrate, z_s_c, 1, S2, m_A, m_B, m_C, m_D, m_E);
			if(Error == 404) return;
			if( Error != -1 ) 
			{
				MessageBox(GetWindow(), FuncErrors(Error));
				return;
			}
			Error = -1;
		}
	}
	if(vecErrs.GetSize()!=0)
	{
		string str = "Были ошибки в следующих книгах: \n";
		for(int aboba=0; aboba<vecErrs.GetSize(); aboba++) str = str + vErrsBooksNames[aboba] + FuncErrors(vecErrs[aboba]) + "\n";
		MessageBox(GetWindow(), str);
	}
	
	////////////////// возвращяю эту хрень к изначальным значениям иначе он ее запоминает и ломается //////////////////
	st_time_end_time[0] = -1;
	st_time_end_time[1] = -1;
	////////////////// возвращяю эту хрень к изначальным значениям иначе он ее запоминает и ломается //////////////////
}

void previous_graph()
{
	Page pg;
	int z_s_c, swap_koeff = -1;
	string pgOrigName;
	GraphLayer graph = Project.ActiveLayer();
    if(!graph)
    {
    	MessageBox(GetWindow(), "Grapg! Смена графика должна осуществляться на окне с графиком");
		return;
    }
    else 
    {
    	pg = graph.GetPage();
		pgOrigName = pg.GetLongName();
		if(pgOrigName.Match("Зонд*")||pgOrigName.Match("Сетка*")||pgOrigName.Match("Цилиндр*")||pgOrigName.Match("Магнит*"))
		{
			if(pgOrigName.Match("Зонд*"))z_s_c=0;
			if(pgOrigName.Match("Сетка*"))z_s_c=1;
			if(pgOrigName.Match("Цилиндр*")||pgOrigName.Match("Магнит*"))z_s_c=2;
		}
		else
		{
			MessageBox(GetWindow(), "!Зонд или !Сетка или !Цилиндр или !Магнит (не тот график)");
			return;
		}
    	
    	if(swap_graph(z_s_c, swap_koeff) == 0) return;
    }
}

void next_graph()
{
	Page pg;
	int z_s_c, swap_koeff = 1;
	string pgOrigName;
	GraphLayer graph = Project.ActiveLayer();
    if(!graph)
    {
    	MessageBox(GetWindow(), "Grapg! Смена графика должна осуществляться на окне с графиком");
		return;
    }
    else 
    {
    	pg = graph.GetPage();
		pgOrigName = pg.GetLongName();
		if(pgOrigName.Match("Зонд*")||pgOrigName.Match("Сетка*")||pgOrigName.Match("Цилиндр*")||pgOrigName.Match("Магнит*"))
		{
			if(pgOrigName.Match("Зонд*"))z_s_c=0;
			if(pgOrigName.Match("Сетка*"))z_s_c=1;
			if(pgOrigName.Match("Цилиндр*")||pgOrigName.Match("Магнит*"))z_s_c=2;
		}
		else
		{
			MessageBox(GetWindow(), "!Зонд или !Сетка или !Цилиндр или !Магнит (не тот график)");
			return;
		}
    	
    	if(swap_graph(z_s_c, swap_koeff) == 0) return;
    }
}

int swap_graph(int z_s_c, int swap_koeff)
{
	Worksheet wks_graph_parent, wks_OrigData, wks_Legend;
	if( Get_OriginData_wks(wks_graph_parent, wks_OrigData, wks_Legend) == 0 ) return 0;
	if(wks_Legend.Columns(0).GetUpperBound() < 3)
	{
		MessageBox(GetWindow(), "Невозможно пересчитать, на графике больше одного отрезка");
		return 0;
	}
	
	if(z_s_c == 0)
	{
		Worksheet wks_sheet1, wks_sheet2,wks_sheet3, wks_sheet4, wks_sheetLegend = wks_Legend.GetPage().Layers(1);
		GraphPage gp = Project.ActiveLayer().GetPage();
		int nPlots = gp.Layers(0).DataPlots.Count(), num_col, index_nach_otrezka;
		is_str_numeric_integer(get_str_from_str_list(wks_Legend.Columns(1).GetLongName(), " ", 1), &index_nach_otrezka);
		wks_sheet1 = WksIdxByName(wks_graph_parent.GetPage(), "OrigData");
		wks_sheet2 = WksIdxByName(wks_graph_parent.GetPage(), "Fit");
		wks_sheet3 = WksIdxByName(wks_graph_parent.GetPage(), "Parameters");
		wks_sheet4 = WksIdxByName(wks_graph_parent.GetPage(), "Filtration");
		if(!wks_sheet1 || !wks_sheet2 || !wks_sheet3)
		{
			MessageBox(GetWindow(), "Ошибка! Не найден лист с данными");
			return 0;
		}
		num_col = ColIdxByUnit(wks_sheet1, (string)index_nach_otrezka).GetIndex();
		if(num_col == 1 && swap_koeff == -1 || num_col == wks_sheet1.GetNumCols() -1 && swap_koeff == 1)
		{
			MessageBox(GetWindow(), "Это крайний график!");
			return 0;
		}
		
		DataRange dr1,dr2,dr3,dr4,dr5; 
		dr1.Add(wks_sheet1, 0, "X");
		dr1.Add(wks_sheet1, num_col+swap_koeff, "Y");
		dr2.Add(wks_sheet2, 0, "X");
		dr2.Add(wks_sheet2, num_col+swap_koeff, "Y");
		if(wks_sheet4 && wks_sheet4.Columns(num_col+swap_koeff).GetUpperBound() > 0) 
		{
			dr3.Add(wks_sheet4, 0, "X");
			dr3.Add(wks_sheet4, num_col+swap_koeff, "Y");
		}
		
		Dataset dsX, dsY, dsSTOLBA, dsA(wks_sheet3.Columns(1)), dsB(wks_sheet3.Columns(2)), dsC(wks_sheet3.Columns(3)), 
			dsD(wks_sheet3.Columns(4)), dsDens(wks_sheet3.Columns(5)), dsPILA(wks_sheet1.Columns(0)), dsTimeSynth(wks_sheet3.Columns(0));
		vector vecA, vecB, vecMinY, vecMaxY, vecPeresechX;
		
		if(dsTimeSynth[num_col+swap_koeff-1] == 0 || dsTimeSynth[num_col+swap_koeff-1] == 0)
		{
			MessageBox(GetWindow(), "Ошибка во времени! Последний отрезок обработан с ошибкой (время начала равно 0). "
				+ "Для решения проблемы удалите отрезки со временем начала 0, или пересчитайте импульс");
			return 0;
		}
		
		for(int ii = 0; ii < nPlots; ii++) gp.Layers(0).RemovePlot(0);
			
		vecA.SetSize(1);
		vecB.SetSize(1);
		vecA[0] = dsA[num_col+swap_koeff-1];
		vecB[0] = dsB[num_col+swap_koeff-1];
		dsSTOLBA.Attach(wks_sheet1.Columns(num_col+swap_koeff));
			
		vecMinY.SetSize(1);
		vecMaxY.SetSize(1);
		vecPeresechX.SetSize(1);
		vecMinY[0] = min(dsSTOLBA);
		vecMaxY[0] = max(dsSTOLBA);
		//vecPeresechX[0]=dens(false, num_col+swap_koeff, dsDens, dsA, dsB, dsC, dsD, dsPILA, dsPILA.GetSize(), 0);
		dsSTOLBA.Attach(wks_sheet2.Columns(num_col+swap_koeff));
		Curve crv(dsPILA,dsSTOLBA);
		vecPeresechX[0] = Curve_xfromY(&crv,0);
		if(vecPeresechX[0] > dsPILA[dsPILA.GetSize()-1]) vecPeresechX[0] = dsPILA[dsPILA.GetSize()-1];
		else if(vecPeresechX[0] < dsPILA[0]) vecPeresechX[0] = dsPILA[0];
			
		dsX.Attach(wks_sheetLegend.Columns(0));
		dsY.Attach(wks_sheetLegend.Columns(1));
		dsX.SetSize(vecA.GetSize()*2);
		dsY.SetSize(vecA.GetSize()*2);
		for(int i=0, j=0; i<vecA.GetSize(); i++, j=j+2)
		{
			dsX[j] = dsPILA[0];
			dsX[j+1] = dsPILA[dsPILA.GetSize()-1];
			dsY[j] = dsX[j]*vecB[i] + vecA[i];
			dsY[j+1] = dsX[j+1]*vecB[i] + vecA[i];
		} 
		dsX.Attach(wks_sheetLegend.Columns(2));
		dsY.Attach(wks_sheetLegend.Columns(3));
		dsX.SetSize(vecMinY.GetSize()*2);
		dsY.SetSize(vecMaxY.GetSize()*2);
		for(int i_1=0, j_1=0; i_1<vecMinY.GetSize(); i_1++, j_1=j_1+2)
		{
			dsX[j_1] = vecPeresechX[i_1];
			dsX[j_1+1] = vecPeresechX[i_1];
			dsY[j_1] = vecMinY[i_1];
			dsY[j_1+1] = vecMaxY[i_1];
		}
		dr4.Add(wks_sheetLegend, 0, "X");
		dr4.Add(wks_sheetLegend, 1, "Y");
		dr5.Add(wks_sheetLegend, 2, "X");
		dr5.Add(wks_sheetLegend, 3, "Y");
		
		GraphLayer gl1 = gp.Layers(0);
		gl1.AddPlot(dr1, IDM_PLOT_LINE);
		gl1.AddPlot(dr2, IDM_PLOT_LINE);
		gl1.AddPlot(dr4, IDM_PLOT_LINE);
		gl1.AddPlot(dr5, IDM_PLOT_LINE);
		if(wks_sheet4 && wks_sheet4.Columns(num_col+swap_koeff).GetUpperBound() > 0) gl1.AddPlot(dr3, IDM_PLOT_LINE);
		
		DataPlot dp1,dp2,dp3,dp4,dp5;
		dp1 = gl1.DataPlots(0);
		dp2 = gl1.DataPlots(1); 
		dp3 = gl1.DataPlots(2);
		dp4 = gl1.DataPlots(3);
		if(wks_sheet4 && wks_sheet4.Columns(num_col+swap_koeff).GetUpperBound() > 0) dp5 = gl1.DataPlots(4);
		
		Tree tr1, tr2, tr3, tr4;
		tr1 = dp1.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
		tr2 = dp2.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
		tr3 = dp3.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
		
		tr1.Root.Line.Width.nVal = 2;
		tr1.Root.Line.Color.nVal = set_color_rnd();
		
		tr2.Root.Line.Width.dVal=2.5;
		tr2.Root.Line.Color.nVal=1;
		
		tr3.Root.Line.Width.dVal=1.1;
		tr3.Root.Line.Color.nVal=3;
		
		dp1.ApplyFormat(tr1, true, true);
		dp2.ApplyFormat(tr2, true, true);
		dp3.ApplyFormat(tr3, true, true);
		dp4.ApplyFormat(tr3, true, true);
		if(wks_sheet4 && wks_sheet4.Columns(num_col+swap_koeff).GetUpperBound() > 0)
		{
			tr4 = dp5.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
			tr4.Root.Line.Width.dVal=2.5;
			tr4.Root.Line.Color.nVal=2;
			dp5.ApplyFormat(tr4, true, true);
			vector<uint> vecUI = {0,4,1,2,3};
			gl1.ReorderPlots(vecUI);
		}
		
		wks_Legend.Columns(0).SetLongName("Участок времени " + wks_sheet1.Columns(num_col+swap_koeff).GetLongName());
		wks_Legend.Columns(1).SetLongName("Строки "+ wks_sheet1.Columns(num_col+swap_koeff).GetUnits());
		wks_Legend.SetCell(0, 1, dsA[num_col+swap_koeff-1]);
		wks_Legend.SetCell(1, 1, dsB[num_col+swap_koeff-1]);
		wks_Legend.SetCell(2, 1, dsC[num_col+swap_koeff-1]);
		wks_Legend.SetCell(3, 1, dsD[num_col+swap_koeff-1]);
		wks_Legend.SetCell(4, 1, dsDens[num_col+swap_koeff-1]);

		gp.SetLongName ( FindPageNameNumber ( get_str_from_str_list(gp.GetLongName(), " ", 0) + " " + get_str_from_str_list(gp.GetLongName(), " ", 1) + " Отрезок " + (num_col+swap_koeff) ) );
		legend_update(gl1);
		
		create_labels(z_s_c);
		
		return -1;
	}
	
	if(z_s_c == 1)
	{
		Worksheet wks_sheet1, wks_sheet2,wks_sheet3, wks_sheet4, wks_sheetLegend = wks_Legend.GetPage().Layers(1);
		GraphPage gp = Project.ActiveLayer().GetPage();
		int nPlots1 = gp.Layers(0).DataPlots.Count(), nPlots2 = gp.Layers(1).DataPlots.Count(), num_col, index_nach_otrezka;
		is_str_numeric_integer(get_str_from_str_list(wks_Legend.Columns(1).GetLongName(), " ", 1), &index_nach_otrezka);
		wks_sheet1 = WksIdxByName(wks_graph_parent.GetPage(), "OrigData");
		wks_sheet2 = WksIdxByName(wks_graph_parent.GetPage(), "Differentiation");
		wks_sheet3 = WksIdxByName(wks_graph_parent.GetPage(), "Parameters");
		wks_sheet4 = WksIdxByName(wks_graph_parent.GetPage(), "Filtration");
		if(!wks_sheet1 || !wks_sheet2 || !wks_sheet3 || !wks_sheet4)
		{
			MessageBox(GetWindow(), "Ошибка! Не найден лист с данными");
			return 0;
		}
		num_col = ColIdxByUnit(wks_sheet1, (string)index_nach_otrezka).GetIndex();
		if(num_col == 1 && swap_koeff == -1 || num_col == wks_sheet1.GetNumCols() -1 && swap_koeff == 1)
		{
			MessageBox(GetWindow(), "Это крайний график!");
			return 0;
		}
		
		DataRange dr1,dr2,dr3,dr4,dr5,dr6,dr7; 
		dr1.Add(wks_sheet1, 0, "X");
		dr1.Add(wks_sheet1, num_col+swap_koeff, "Y");
		dr2.Add(wks_sheet4, 0, "X");
		dr2.Add(wks_sheet4, num_col+swap_koeff, "Y");
		dr3.Add(wks_sheet2, 0, "X");
		dr3.Add(wks_sheet2, num_col+swap_koeff, "Y");
		
		Dataset ds_maxVecDiff_X, ds_maxVecDiff_Y, dsSTOLBA, ds_Polyvisota, ds_index_sred_lev, ds_index_sred_prav, ds_Y_lev, ds_Y_prav, 
			dsPILA(wks_sheet1.Columns(0)), dsMaxValue(wks_sheet3.Columns(1)), dsTemp, dsTimeSynth(wks_sheet3.Columns(0));
		vector vec_maxVecDiff_Y, vec_maxVecDiff_X, vec_index_sred_lev, vec_index_sred_prav, vec_Polyvisota, vecY_lev, vecY_prav;
		
		if(dsTimeSynth[num_col+swap_koeff-1] == 0 || dsTimeSynth[num_col+swap_koeff-1] == 0)
		{
			MessageBox(GetWindow(), "Ошибка во времени! Последний отрезок обработан с ошибкой (время начала равно 0). "
				+ "Для решения проблемы удалите отрезки со временем начала 0, или пересчитайте импульс");
			return 0;
		}
		
		for(int ii = 0; ii < nPlots1; ii++) gp.Layers(0).RemovePlot(0);
		for(int jj = 0; jj < nPlots2; jj++) gp.Layers(1).RemovePlot(0);
		
		vec_maxVecDiff_Y.SetSize(2);
		vec_maxVecDiff_X.SetSize(2);
		vec_index_sred_lev.SetSize(2);
		vec_index_sred_prav.SetSize(2);
		
		vector v_whm_w_peak(10);
		/*
		вектор v_whm_w_peak:
			whm
			-[0] ширина на полувысоте
			-[1] Y на полувысоте
			-[2] X на полувысоте слева
			-[3] X на полувысоте справа
			W
			-[4] значение коэффициента W
			-[5] Y при W
			-[6] X при W слева
			-[7] X при W справа
			Peak
			-[8] Y пика
			-[9] X пика
		*/
		
		vec_maxVecDiff_Y[0] = vec_maxVecDiff_Y[1] = dsMaxValue[num_col+swap_koeff-1];
		vec_maxVecDiff_X[0] = min(dsPILA);
		vec_maxVecDiff_X[1] = max(dsPILA);
			
		dsSTOLBA.Attach(wks_sheet2.Columns(num_col+swap_koeff));
		
		vec_Polyvisota.SetSize(2);
		vecY_lev.SetSize(2);
		vecY_prav.SetSize(2);
		
		vecY_lev[0] = vecY_prav[0] = min(dsSTOLBA);
		vecY_lev[1] = vecY_prav[1] = max(dsSTOLBA);
			
		v_whm_w_peak = find_whm_Wparam_LRmidindxs(dsPILA, dsSTOLBA, 0, 0, -1);
		
		vec_Polyvisota[0] = vec_Polyvisota[1] = v_whm_w_peak[1];
		vec_index_sred_lev[0] = vec_index_sred_lev[1] = v_whm_w_peak[2];
		vec_index_sred_prav[0] = vec_index_sred_prav[1] = v_whm_w_peak[3];
			
		ds_maxVecDiff_X.Attach(wks_sheetLegend.Columns(0));
		ds_maxVecDiff_Y.Attach(wks_sheetLegend.Columns(1));
		ds_Polyvisota.Attach(wks_sheetLegend.Columns(2));
		ds_index_sred_lev.Attach(wks_sheetLegend.Columns(3));
		ds_Y_lev.Attach(wks_sheetLegend.Columns(4));
		ds_index_sred_prav.Attach(wks_sheetLegend.Columns(5));
		ds_Y_prav.Attach(wks_sheetLegend.Columns(6));
		
		ds_maxVecDiff_X=vec_maxVecDiff_X;
		ds_maxVecDiff_Y=vec_maxVecDiff_Y;
		ds_Polyvisota=vec_Polyvisota;
		ds_index_sred_lev=vec_index_sred_lev;
		ds_Y_lev=vecY_lev;
		ds_index_sred_prav=vec_index_sred_prav;
		ds_Y_prav=vecY_prav;
		
		dr4.Add(wks_sheetLegend, 0, "X");
		dr4.Add(wks_sheetLegend, 1, "Y");
		dr5.Add(wks_sheetLegend, 0, "X");
		dr5.Add(wks_sheetLegend, 2, "Y");
		dr6.Add(wks_sheetLegend, 3, "X");
		dr6.Add(wks_sheetLegend, 4, "Y");
		dr7.Add(wks_sheetLegend, 5, "X");
		dr7.Add(wks_sheetLegend, 6, "Y");
		
		GraphLayer gl1 = gp.Layers(0), gl2 = gp.Layers(1);
		gl1.AddPlot(dr1, IDM_PLOT_LINE);
		gl1.AddPlot(dr2, IDM_PLOT_LINE);
		gl2.AddPlot(dr4, IDM_PLOT_LINE);
		gl2.AddPlot(dr5, IDM_PLOT_LINE);
		gl2.AddPlot(dr6, IDM_PLOT_LINE);
		gl2.AddPlot(dr7, IDM_PLOT_LINE);
		gl2.AddPlot(dr3, IDM_PLOT_LINE);
		
		DataPlot dp1,dp2,dp3,dp4,dp5,dp6,dp7;
		dp1 = gl1.DataPlots(0);
		dp2 = gl1.DataPlots(1);
		dp3 = gl2.DataPlots(0);
		dp4 = gl2.DataPlots(1);
		dp5 = gl2.DataPlots(2);
		dp6 = gl2.DataPlots(3);
		dp7 = gl2.DataPlots(4);

		Tree tr1, tr2, tr3, tr6;
		tr1 = dp1.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
		tr2 = dp2.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
		tr3 = dp3.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
		tr6 = dp6.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
		
		tr1.Root.Line.Width.nVal = 2;
		tr1.Root.Line.Color.nVal = set_color_rnd();
		
		tr2.Root.Line.Width.dVal=2.5;
		tr2.Root.Line.Color.nVal=1;
		
		tr3.Root.Line.Width.dVal=1.1;
		tr3.Root.Line.Color.nVal=3;

		tr6.Root.Line.Width.dVal=2.5;
		tr6.Root.Line.Color.nVal=2;
		
		dp1.ApplyFormat(tr1, true, true);
		dp2.ApplyFormat(tr2, true, true);
		dp3.ApplyFormat(tr3, true, true);
		dp4.ApplyFormat(tr3, true, true);
		dp5.ApplyFormat(tr3, true, true);
		dp6.ApplyFormat(tr3, true, true);
		dp7.ApplyFormat(tr6, true, true);
		
		wks_Legend.Columns(0).SetLongName("Участок времени " + wks_sheet1.Columns(num_col+swap_koeff).GetLongName());
		wks_Legend.Columns(1).SetLongName("Строки "+ wks_sheet1.Columns(num_col+swap_koeff).GetUnits());

		dsTemp.Attach(wks_sheet3.Columns(2));
		wks_Legend.SetCell(0, 1, dsTemp[num_col+swap_koeff-1]);
		dsTemp.Attach(wks_sheet3.Columns(3));
		wks_Legend.SetCell(1, 1, dsTemp[num_col+swap_koeff-1]);

		gp.SetLongName ( FindPageNameNumber ( get_str_from_str_list(gp.GetLongName(), " ", 0) + " " + get_str_from_str_list(gp.GetLongName(), " ", 1) + " Отрезок " + (num_col+swap_koeff) ) );
		legend_update(gl1);
		legend_combine(gp);
		
		create_labels(z_s_c);
		
		return -1;
	}
	
	if(z_s_c == 2)
	{
		Worksheet wks_sheet1, wks_sheet2,wks_sheet3, wks_sheet4, wks_sheetLegend = wks_Legend.GetPage().Layers(1);
		GraphPage gp = Project.ActiveLayer().GetPage();
		int nPlots = gp.Layers(0).DataPlots.Count(), num_col, index_nach_otrezka;
		is_str_numeric_integer(get_str_from_str_list(wks_Legend.Columns(1).GetLongName(), " ", 1), &index_nach_otrezka);
		wks_sheet1 = WksIdxByName(wks_graph_parent.GetPage(), "OrigData");
		wks_sheet2 = WksIdxByName(wks_graph_parent.GetPage(), "Fit");
		wks_sheet3 = WksIdxByName(wks_graph_parent.GetPage(), "Parameters");
		wks_sheet4 = WksIdxByName(wks_graph_parent.GetPage(), "Filtration");
		if(!wks_sheet1 || !wks_sheet2 || !wks_sheet3)
		{
			MessageBox(GetWindow(), "Ошибка! Не найден лист с данными");
			return 0;
		}
		num_col = ColIdxByUnit(wks_sheet1, (string)index_nach_otrezka).GetIndex();
		if(num_col == 1 && swap_koeff == -1 || num_col == wks_sheet1.GetNumCols() -1 && swap_koeff == 1)
		{
			MessageBox(GetWindow(), "Это крайний график!");
			return 0;
		}
		
		DataRange dr1,dr2,dr3,dr4,dr5,dr6,dr7; 
		dr1.Add(wks_sheet1, 0, "X");
		dr1.Add(wks_sheet1, num_col+swap_koeff, "Y");
		dr2.Add(wks_sheet2, 0, "X");
		dr2.Add(wks_sheet2, num_col+swap_koeff, "Y");
		if(wks_sheet4 && wks_sheet4.Columns(num_col+swap_koeff).GetUpperBound() > 0) 
		{
			dr3.Add(wks_sheet4, 0, "X");
			dr3.Add(wks_sheet4, num_col+swap_koeff, "Y");
		}
		
		Dataset ds_maxVecDiff_X, ds_maxVecDiff_Y, dsSTOLBA, ds_Polyvisota, ds_index_sred_lev, ds_index_sred_prav, ds_Y_lev, ds_Y_prav, dsPILA(wks_sheet1.Columns(0)), 
			dsMaxValue(wks_sheet3.Columns(1)), dsTemp(wks_sheet3.Columns(2)), dsPeakVoltage(wks_sheet3.Columns(3)), dsTimeSynth(wks_sheet3.Columns(0));
		vector vec_maxVecDiff_Y, vec_maxVecDiff_X, vec_index_sred_lev, vec_index_sred_prav, vec_Polyvisota, vecY_lev, vecY_prav;
		
		if(dsTimeSynth[num_col+swap_koeff-1] == 0 || dsTimeSynth[num_col+swap_koeff-1] == 0)
		{
			MessageBox(GetWindow(), "Ошибка во времени! Последний отрезок обработан с ошибкой (время начала равно 0). "
				+ "Для решения проблемы удалите отрезки со временем начала 0, или пересчитайте импульс");
			return 0;
		}
		
		for(int ii = 0; ii < nPlots; ii++) gp.Layers(0).RemovePlot(0);
		
		vec_maxVecDiff_Y.SetSize(2);
		vec_maxVecDiff_X.SetSize(2);
		vec_index_sred_lev.SetSize(2);
		vec_index_sred_prav.SetSize(2);
		
		vector v_whm_w_peak(10);
		/*
		вектор v_whm_w_peak:
			whm
			-[0] ширина на полувысоте
			-[1] Y на полувысоте
			-[2] X на полувысоте слева
			-[3] X на полувысоте справа
			W
			-[4] значение коэффициента W
			-[5] Y при W
			-[6] X при W слева
			-[7] X при W справа
			Peak
			-[8] Y пика
			-[9] X пика
		*/
		
		vec_maxVecDiff_Y[0] = vec_maxVecDiff_Y[1] = dsMaxValue[num_col+swap_koeff-1];
		vec_maxVecDiff_X[0] = min(dsPILA);
		vec_maxVecDiff_X[1] = max(dsPILA);
			
		dsSTOLBA.Attach(wks_sheet2.Columns(num_col+swap_koeff));
		
		vec_Polyvisota.SetSize(2);
		vecY_lev.SetSize(2);
		vecY_prav.SetSize(2);
		
		vecY_lev[0] = vecY_prav[0] = min(dsSTOLBA);
		vecY_lev[1] = vecY_prav[1] = max(dsSTOLBA);
			
		v_whm_w_peak = find_whm_Wparam_LRmidindxs(dsPILA, dsSTOLBA, 1, dsTemp[num_col+swap_koeff-1], -1);
		
		vec_Polyvisota[0] = vec_Polyvisota[1] = v_whm_w_peak[5];
		vec_index_sred_lev[0] = vec_index_sred_lev[1] = v_whm_w_peak[6];
		vec_index_sred_prav[0] = vec_index_sred_prav[1] = v_whm_w_peak[7];
		
		ds_maxVecDiff_X.Attach(wks_sheetLegend.Columns(0));
		ds_maxVecDiff_Y.Attach(wks_sheetLegend.Columns(1));
		ds_Polyvisota.Attach(wks_sheetLegend.Columns(2));
		ds_index_sred_lev.Attach(wks_sheetLegend.Columns(3));
		ds_Y_lev.Attach(wks_sheetLegend.Columns(4));
		ds_index_sred_prav.Attach(wks_sheetLegend.Columns(5));
		ds_Y_prav.Attach(wks_sheetLegend.Columns(6));
		
		ds_maxVecDiff_X=vec_maxVecDiff_X;
		ds_maxVecDiff_Y=vec_maxVecDiff_Y;
		ds_Polyvisota=vec_Polyvisota;
		ds_index_sred_lev=vec_index_sred_lev;
		ds_Y_lev=vecY_lev;
		ds_index_sred_prav=vec_index_sred_prav;
		ds_Y_prav=vecY_prav;
		
		dr4.Add(wks_sheetLegend, 0, "X");
		dr4.Add(wks_sheetLegend, 1, "Y");
		dr5.Add(wks_sheetLegend, 0, "X");
		dr5.Add(wks_sheetLegend, 2, "Y");
		dr6.Add(wks_sheetLegend, 3, "X");
		dr6.Add(wks_sheetLegend, 4, "Y");
		dr7.Add(wks_sheetLegend, 5, "X");
		dr7.Add(wks_sheetLegend, 6, "Y");
		
		GraphLayer gl1 = gp.Layers(0);
		gl1.AddPlot(dr1, IDM_PLOT_LINE);
		if(wks_sheet4 && wks_sheet4.Columns(num_col+swap_koeff).GetUpperBound() > 0) gl1.AddPlot(dr3, IDM_PLOT_LINE);
		gl1.AddPlot(dr4, IDM_PLOT_LINE);
		gl1.AddPlot(dr5, IDM_PLOT_LINE);
		gl1.AddPlot(dr6, IDM_PLOT_LINE);
		gl1.AddPlot(dr7, IDM_PLOT_LINE);
		gl1.AddPlot(dr2, IDM_PLOT_LINE);
		
		DataPlot dp1,dp2,dp3,dp4,dp5,dp6,dp7;
		dp1 = gl1.DataPlots(0);
		if(wks_sheet4 && wks_sheet4.Columns(num_col+swap_koeff).GetUpperBound() > 0)
		{
			dp2 = gl1.DataPlots(6);
			dp3 = gl1.DataPlots(2);
			dp7 = gl1.DataPlots(3);
			dp4 = gl1.DataPlots(4);
			dp5 = gl1.DataPlots(5);
			dp6 = gl1.DataPlots(1);
		}
		else
		{
			dp2 = gl1.DataPlots(5);
			dp3 = gl1.DataPlots(1);
			dp7 = gl1.DataPlots(2);
			dp4 = gl1.DataPlots(3);
			dp5 = gl1.DataPlots(4);
		}

		Tree tr1, tr2, tr3, tr6;	
		tr1 = dp1.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
		tr2 = dp2.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
		tr3 = dp3.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
		if(wks_sheet4 && wks_sheet4.Columns(num_col+swap_koeff).GetUpperBound() > 0) tr6 = dp6.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
		
		tr1.Root.Line.Width.nVal = 2;
		tr1.Root.Line.Color.nVal = set_color_rnd();
		
		tr2.Root.Line.Width.dVal=2.5;
		tr2.Root.Line.Color.nVal=1;
		
		tr3.Root.Line.Width.dVal=1.1;
		tr3.Root.Line.Color.nVal=3;

		tr6.Root.Line.Width.dVal=2.5;
		tr6.Root.Line.Color.nVal=2;
		
		dp1.ApplyFormat(tr1, true, true);
		dp2.ApplyFormat(tr2, true, true);
		dp3.ApplyFormat(tr3, true, true);
		dp4.ApplyFormat(tr3, true, true);
		dp5.ApplyFormat(tr3, true, true);
		if(wks_sheet4 && wks_sheet4.Columns(num_col+swap_koeff).GetUpperBound() > 0) dp6.ApplyFormat(tr6, true, true);
		dp7.ApplyFormat(tr3, true, true);
		
		wks_Legend.Columns(0).SetLongName("Участок времени " + wks_sheet1.Columns(num_col+swap_koeff).GetLongName());
		wks_Legend.Columns(1).SetLongName("Строки "+ wks_sheet1.Columns(num_col+swap_koeff).GetUnits());

		dsTemp.Attach(wks_sheet3.Columns(2));
		wks_Legend.SetCell(0, 1, dsTemp[num_col+swap_koeff-1]);
		dsTemp.Attach(wks_sheet3.Columns(3));
		wks_Legend.SetCell(1, 1, dsTemp[num_col+swap_koeff-1]);

		gp.SetLongName ( FindPageNameNumber ( get_str_from_str_list(gp.GetLongName(), " ", 0) + " " + get_str_from_str_list(gp.GetLongName(), " ", 1) + " Отрезок " + (num_col+swap_koeff) ) );
		legend_update(gl1);
		
		create_labels(z_s_c);
		
		return -1;
	}
	
	return -1;
}

void interf_transform()
{
	Worksheet wks = Project.ActiveLayer(), wks_slave;
	if(!wks) 
    {
    	MessageBox(GetWindow(), "wks!");
    	return;
    }
	if(wks.GetPage().GetLongName().Match("Зонд*") || wks.GetPage().GetLongName().Match("Сетка*")
		|| wks.GetPage().GetLongName().Match("Цилиндр*") || wks.GetPage().GetLongName().Match("Магнит*"))
	{
		MessageBox(GetWindow(), "Не тот лист! Выберите лист с результатами эксперимента");
		return;
	}
    vector<string> vBooksNames, vErrsBooksNames;
	foreach(WorksheetPage pg in Project.WorksheetPages)
	{
		int BeginVal;
		is_str_numeric_integer(get_str_from_str_list(pg.GetLongName(), "_", 0), &BeginVal);
		if(BeginVal>10000) vBooksNames.Add(pg.GetLongName());
	}
	
	GETN_BOX( treeTest );
    GETN_NUM(decayT1, "Частота дискретизации интерферометра, Гц", ch_dsc_ntrfrmtr)
    GETN_RADIO_INDEX_EX(decayT2, "Выровнять столбец по остальным данным", 0, "Да|Нет")
    
    GETN_BEGIN_BRANCH(Details1, "Фильтрация данных") GETN_CHECKBOX_BRANCH(1)
    GETN_OPTION_BRANCH(GETNBRANCH_OPEN)
    GETN_SLIDEREDIT(decayT14, "Часть точек фильтрации", 0.0001, "0.0001|0.01|98") 
    GETN_END_BRANCH(Details1)
    
    GETN_BEGIN_BRANCH(Details, "Нужно обработать несколько книг") GETN_CHECKBOX_BRANCH(0)
    for(int i_1 = 0; i_1<vBooksNames.GetSize(); i_1++) 
    {
		GETN_CHECK(NodeName, vBooksNames[i_1], true)
    }
    GETN_END_BRANCH(Details)
    
    if(0==GetNBox( treeTest, "Обработка данных с зонда" )) return;
    
    time_t start_time, finish_time;
	time(&start_time);
    
    double Chast_to_Filtrate;
    int vyr_or_not, UseSameName_or_Not_NTRF, True_or_False, Error = -1;
    string ReportString_NTRF;
    vector vecErrs;
    TreeNode tn1;
 
	treeTest.GetNode("decayT1").GetValue(ch_dsc_ntrfrmtr);
	treeTest.GetNode("decayT2").GetValue(vyr_or_not);
	if(treeTest.Details1.Use)
	{
		treeTest.GetNode("Details1").GetNode("decayT14").GetValue(Chast_to_Filtrate);
		Chast_to_Filtrate = Chast_to_Filtrate/2;
	}
	else Chast_to_Filtrate = 0;
   	
   	Column colNTRF = ColIdxByName(wks, s_ntrfrmtr);
	if(!colNTRF)
	{
		GETN_BOX( treeTest1 );
		GETN_STR(STR, "!Не найден столбец с именем " + "\"" + s_ntrfrmtr + "\"" + "!", "") GETN_INFO
		GETN_INTERACTIVE(interactive, "Выберите столбец с нужными данными", "[Book1]Sheet1!B")
		GETN_OPTION_INTERACTIVE_CONTROL(ICOPT_RESTRICT_TO_ONE_DATA) 
		if(treeTest.Details.Use)
		{
			GETN_RADIO_INDEX_EX(decayT1, "Использовать колонку с таким же именем для расчета остальных книг", 0, "Да|Нет")
		}
		if(0==GetNBox( treeTest1, "Ошибка!" )) return;
		
		treeTest1.GetNode("interactive").GetValue(ReportString_NTRF);
		treeTest1.GetNode("decayT1").GetValue(UseSameName_or_Not_NTRF);
		ReportString_NTRF = get_str_from_str_list(ReportString_NTRF, "\"", 1);
		colNTRF = ColIdxByName(wks, ReportString_NTRF);
		if(!colNTRF)
		{
			MessageBox(GetWindow(), "Ошибка. Возможно, выбранная вами колонка пустая");
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
							vErrsBooksNames.Add("Интерферометр "+vBooksNames[i_2]+": ");
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
					Error = interf_prog(colNTRF_slave, ch_dsc_ntrfrmtr, wks_slave, Chast_to_Filtrate, vyr_or_not);
					if(Error == 404) return;
					if(Error != -1) 
					{
						vecErrs.Add(Error);
						vErrsBooksNames.Add("Интерферометр "+vBooksNames[i_2]+": ");
					}
					Error = -1;
				}
			}
		}
	}
	else 
	{
		Error = interf_prog(colNTRF, ch_dsc_ntrfrmtr, wks, Chast_to_Filtrate, vyr_or_not);
		if(Error == 404) return;
		if( Error != -1 ) 
		{
			MessageBox(GetWindow(), FuncErrors(Error));
			return;
		}
	}
	if(vecErrs.GetSize()!=0)
	{
		string str = "Были ошибки в следующих книгах: \n";
		for(int aboba=0; aboba<vecErrs.GetSize(); aboba++) str = str + vErrsBooksNames[aboba] + FuncErrors(vecErrs[aboba]) + "\n";
		MessageBox(GetWindow(), str);
	}
	
	time(&finish_time);
	printf("Run time all prog %d secs\n", finish_time - start_time);
}

int interf_prog(Column colNTRF, int ch_dsc_ntrfrmtr, Worksheet wks, double Chast_to_Filtrate, int vyr_or_not)
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
	if(Chast_to_Filtrate != 0) 
    	ocmath_savitsky_golay(N, N_posrednik, N.GetSize(), N.GetSize()*Chast_to_Filtrate, -1, 5, 0, EDGEPAD_REFLECT);
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
	
	//// массив времени
	//double dt = tmax/(double)k;
	//for(i = 0; i < k; i++)
		//Time[i] = dt*i;  
		//
	//Worksheet wks_test;
	//wks_test.Create("Origin");
	//
	//Dataset time(wks_test.Columns(0)), data(wks_test.Columns(1));
	//time = Time;
	//data = N;
	
	return -1;
}

int auto_chdsc(double &chastota_discret_pily, double &chastota_discret_zonda, Column colPILA, Column colTOK, int ch_pil, int z_s_c, Worksheet wks, int diagnostic_num)
{
	time_t start_time, finish_time;
	time(&start_time);
	
	switch(z_s_c)
	{
		case 0:
			if(!colPILA) colPILA = ColIdxByName(wks, s_pz);	
			if(diagnostic_num == 0 && !colTOK) colTOK = ColIdxByName(wks, s_tz);
			if(diagnostic_num == 1 && !colTOK) colTOK = ColIdxByName(wks, s_tz2);
			break;
		case 1:
			if(!colPILA) colPILA = ColIdxByName(wks, s_ps);	
			if(diagnostic_num == 0 && !colTOK) colTOK = ColIdxByName(wks, s_tst);
			if(diagnostic_num == 1 && !colTOK) colTOK = ColIdxByName(wks, s_tssh);
			break;
		case 2:
			if(diagnostic_num == 0 && !colPILA) colPILA = ColIdxByName(wks, s_pc);
			if(diagnostic_num == 1 && !colPILA) colPILA = ColIdxByName(wks, s_pm);
			if(diagnostic_num == 0 && !colTOK) colTOK = ColIdxByName(wks, s_tc);
			if(diagnostic_num == 1 && !colTOK) colTOK = ColIdxByName(wks, s_tm);
			break;
		case -1:
			break;
	}

	if(!colTOK) return 222;
	if(!colPILA) return 333;
	
	vector vecPILA(colPILA), vecTOK(colTOK), vecX(vecPILA.GetSize()/10), vec_peaks(4), vec_slave(vecPILA.GetSize()/10), vecPILA_posrednik(colPILA);
	int i,k;
	
	//////////////////СЛАБО ФИЛЬТРУЕМ ДЛЯ ТОГО ЧТОБЫ УБРАТЬ ВСПЛЕСКИ//////////////////
	ocmath_savitsky_golay(vecPILA, vecPILA_posrednik, vecPILA.GetSize(), vecPILA.GetSize()/10000, -1, 1, 0, EDGEPAD_REFLECT);
	vecPILA = vecPILA_posrednik;
	//////////////////СЛАБО ФИЛЬТРУЕМ ДЛЯ ТОГО ЧТОБЫ УБРАТЬ ВСПЛЕСКИ//////////////////
	
	for(i=0;i<vecX.GetSize();i++) {
		vecX[i]=i;
		vec_slave[i]=vecPILA[i];
	}
	
	Curve crv(vecX,vec_slave);
	
	//////////////////////////////////////////////
	/*
	vec_peaks(размер - 4): первые два значения - индексы пиков, вторые 2 - их значения
	*/
	//////////////////////////////////////////////
	
	vec_peaks[0] = xatymax(crv);
	vec_peaks[2] = max(vec_slave);
	
	vec_slave.SetSize( vec_slave.GetSize()/2 );
	vecX.SetSize( vec_slave.GetSize() );
	if(vec_peaks[0] <= vec_slave.GetSize()) 
	{
		for(i = vec_slave.GetSize() + 1, k = 0; k < vec_slave.GetSize(); i++, k++)
		{
			vecX[k] = i;
			vec_slave[k] = vecPILA[i];
		} 
	}
	if( (vec_peaks[0] > vec_slave.GetSize()) && (vec_peaks[0] <= vec_slave.GetSize()*2) ) 
	{
		for(i = 0, k = 0; k < vec_slave.GetSize(); i++, k++)
		{
			vecX[k] = i;
			vec_slave[k] = vecPILA[i];
		}
	}

	Curve crv1(vecX,vec_slave);
	
	vec_peaks[1] = xatymax(crv1);
	vec_peaks[3] = max(vec_slave);
	
	i=0;
	while(find_peaks(vecX, vec_slave, vec_peaks, vecPILA)) {
		i++;
		if(i>10) return 444;
	}
	
	/* определяем частоту дискретизации пилы */
	int chdsc_orig = abs(vec_peaks[1]-vec_peaks[0])*ch_pil, todel = (int)abs(vec_peaks[1]-vec_peaks[0])*ch_pil % koeff_rounding;
	if((double)todel/koeff_rounding <= 0.5) chastota_discret_pily = chdsc_orig - todel;
	else chastota_discret_pily = chdsc_orig + koeff_rounding - todel;
	
	/* определяем частоту дискретизации тока */
	if(vecPILA.GetSize() == vecTOK.GetSize()) chastota_discret_zonda = chastota_discret_pily;
	else chastota_discret_zonda = (int)(chastota_discret_pily / ((double)vecPILA.GetSize() / (double)vecTOK.GetSize()));
	
	/* определяем длину участка пилы */
	todel = (int)abs(vec_peaks[1]-vec_peaks[0]) % 10;
	if((double)todel/10 <= 0.5) NumPointsOfOriginalPila = abs(vec_peaks[1]-vec_peaks[0]) - todel;
	else NumPointsOfOriginalPila = abs(vec_peaks[1]-vec_peaks[0]) + 10 - todel;
	
	/* округляем частоту дискретизации тока */
	chdsc_orig = chastota_discret_zonda;
	todel = (int)chastota_discret_zonda % koeff_rounding;
	if((double)todel/koeff_rounding <= 0.5) chastota_discret_zonda = chdsc_orig - todel;
	else chastota_discret_zonda = chdsc_orig + koeff_rounding - todel;
	
	/* проверка ошибок */
	if(chastota_discret_pily == 0 || chastota_discret_zonda == 0 || chastota_discret_pily == NANUM || chastota_discret_zonda == NANUM) return 777;
	
	time(&finish_time);
	printf("Run time auto_chdsc %d secs\n", finish_time - start_time);
	
	return -1;
}

bool find_peaks(vector vecX, vector &vec_slave, vector &vec_peaks, vector vecPILA)
{
	int i, k, size_orig = abs(vec_peaks[1]-vec_peaks[0]);
	
	vec_slave.SetSize( size_orig*0.4 );
	vecX.SetSize( vec_slave.GetSize() );
	
	/* переписываем уменьшенный vecX */
	if(vec_peaks[0] <= vec_peaks[1]) 
		for(i = vec_peaks[0] + size_orig*0.3, k = 0; k < vecX.GetSize(); i++, k++) 
			vecX[k] = i;
	else 
		for(i = vec_peaks[1] + size_orig*0.3, k = 0; k < vecX.GetSize(); i++, k++) 
			vecX[k] = i;
	/* переписываем уменьшенный вектор пилы */
	for(i = vec_peaks[0] + size_orig*0.3, k = 0; k < vec_slave.GetSize(); i++, k++) 
		vec_slave[k] = vecPILA[i];
	
	if(abs(vec_peaks[2]) >= abs(vec_peaks[3]))
	{
		if( abs(max(vec_slave)/vec_peaks[2]) > 1-razbros_po_pile && abs(max(vec_slave)/vec_peaks[2]) < 1+razbros_po_pile ) {
			Curve crv(vecX,vec_slave);
			vec_peaks[0] = xatymax(crv);
			vec_peaks[2] = max(vec_slave);
			return true;
		}
		else return false;
	}
	else 
	{
		if( abs(max(vec_slave)/vec_peaks[3]) > 1-razbros_po_pile && abs(max(vec_slave)/vec_peaks[3]) < 1+razbros_po_pile ) {
			Curve crv(vecX,vec_slave);
			vec_peaks[1] = xatymax(crv);
			vec_peaks[3] = max(vec_slave);
			return true;
		}
		else return false;
	}
}

int create_wks_z_s_c(int z_s_c, Page &pg, int SVCH_or_VCH, int He_or_Ar, int kolvo_otrezkov, vector &vec_sokrashPILA, double Chast_to_Filtrate, int resistance, 
					int chastota_discret_pily, int chastota_discret_zonda, vector &vec_Otr_Start_Indxs, int chastota_pily, int FirstAnalyse_or_Second, 
					double Chast_to_Filtrate_Diff, int Poly_Order, string pgOrigName)
{
	Worksheet wks_test, wks_test_1, wks_test_2, wks_test_3;
	string newPgName, wksUnits, str_SVCH_or_VCH, str_He_or_Ar, strFilt, otsecheniya;
	
	switch(SVCH_or_VCH)
	{
		case 0:
			str_SVCH_or_VCH = svch;
			break;
		case 1:
			str_SVCH_or_VCH = vch;
			break;
		case 2:
			str_SVCH_or_VCH = icr;
			break;
	}
	
	if(He_or_Ar==0) str_He_or_Ar = "He";
	else str_He_or_Ar = "Ar";
	
	wks_test = pg.Layers(0);
	while( wks_test.DeleteCol(0) );
	if(kolvo_otrezkov>1)wks_test.SetSize(vec_sokrashPILA.GetSize(), kolvo_otrezkov);
	else wks_test.SetSize(vec_sokrashPILA.GetSize(), 2);
	wks_test.Columns(0).SetLongName("pila");
	wks_test.SetName("OrigData");
	
	if(z_s_c == 1)
	{
		pg.AddLayer("Filtration");
		wks_test_1 = pg.Layers(1);
		while( wks_test_1.DeleteCol(0) );
		if(kolvo_otrezkov>1) wks_test_1.SetSize(vec_sokrashPILA.GetSize(), kolvo_otrezkov);
		else wks_test_1.SetSize(vec_sokrashPILA.GetSize(), 2);
		wks_test_1.Columns(0).SetLongName("pila");
		
		pg.AddLayer("Differentiation");
		wks_test_2 = pg.Layers(2);
		while( wks_test_2.DeleteCol(0) );
		if(kolvo_otrezkov>1) wks_test_2.SetSize(vec_sokrashPILA.GetSize(), kolvo_otrezkov);
		else wks_test_2.SetSize(vec_sokrashPILA.GetSize(), 2);
		wks_test_2.Columns(0).SetLongName("pila");
		
		pg.AddLayer("Parameters");
		wks_test_3 = pg.Layers(3);
		while( wks_test_3.DeleteCol(0) );
		wks_test_3.SetSize(kolvo_otrezkov-1, 4);
		
		if(Chast_to_Filtrate_Diff == 0) strFilt.Format("%.2f", Chast_to_Filtrate*2);
		else strFilt.Format("%.2f Pts %.2f PO %i", Chast_to_Filtrate*2, Chast_to_Filtrate_Diff*2, Poly_Order);
	}
	else
	{	
		pg.AddLayer("Fit");
		wks_test_1 = pg.Layers(1);
		while( wks_test_1.DeleteCol(0) );
		if(kolvo_otrezkov>1)wks_test_1.SetSize(vec_sokrashPILA.GetSize(), kolvo_otrezkov);
		else wks_test_1.SetSize(vec_sokrashPILA.GetSize(), 2);
		wks_test_1.Columns(0).SetLongName("pila");
		
		pg.AddLayer("Parameters");
		wks_test_2 = pg.Layers(2);
		while( wks_test_2.DeleteCol(0) );
		if(z_s_c==0) wks_test_2.SetSize(kolvo_otrezkov-1, 6);
		else wks_test_2.SetSize(kolvo_otrezkov-1, 5);
		
		if(Chast_to_Filtrate != 0)
		{
			pg.AddLayer("Filtration");
			wks_test_3 = pg.Layers(3);
			while( wks_test_3.DeleteCol(0) );
			if(kolvo_otrezkov>1)wks_test_3.SetSize(vec_sokrashPILA.GetSize(), kolvo_otrezkov);
			else wks_test_3.SetSize(vec_sokrashPILA.GetSize(), 2);
			wks_test_3.Columns(0).SetLongName("pila");
		}
	}
	
	wks_test.Columns(0).SetComments("0");
	wks_test_1.Columns(0).SetComments("0");
	if(z_s_c == 1) wks_test_2.Columns(0).SetComments("0");
	wks_test.Columns(0).SetType(OKDATAOBJ_DESIGNATION_X);
	wks_test_1.Columns(0).SetType(OKDATAOBJ_DESIGNATION_X);
	wks_test_2.Columns(0).SetType(OKDATAOBJ_DESIGNATION_X);
	if(z_s_c == 1) wks_test_3.Columns(0).SetType(OKDATAOBJ_DESIGNATION_X);
	
	wksUnits.Format("%s %s %i", str_SVCH_or_VCH, str_He_or_Ar, (int)resistance);
	wks_test.Columns(0).SetUnits(wksUnits);
	wks_test_1.Columns(0).SetUnits(wksUnits);
	if(z_s_c == 1) wks_test_2.Columns(0).SetUnits(wksUnits);
	
	Column col_sokrashPILA = wks_test.Columns(0), colTIME_synth=wks_test_2.Columns(0), colA=wks_test_2.Columns(1), colB=wks_test_2.Columns(2), colC=wks_test_2.Columns(3), colD=wks_test_2.Columns(4), colDens;
	
	if(z_s_c == 1)
	{
		wks_test_3.Columns(0).SetLongName("time");
		wks_test_3.Columns(0).SetUnits("chpil " + (int)chastota_discret_pily);
		wks_test_3.Columns(0).SetComments("chtok " + (int)chastota_discret_zonda);
	}
	else
	{
		colTIME_synth.SetLongName("time");
		colTIME_synth.SetUnits("chpil " + (int)chastota_discret_pily);
		colTIME_synth.SetComments("chtok " + (int)chastota_discret_zonda);
	}
	
	switch(z_s_c)
	{
		case 0:
			if(FirstAnalyse_or_Second == 0) newPgName.Format("Зонд1 %s %s %s", str_SVCH_or_VCH, str_He_or_Ar, pgOrigName);
			if(FirstAnalyse_or_Second == 1) newPgName.Format("Зонд2 %s %s %s", str_SVCH_or_VCH, str_He_or_Ar, pgOrigName);;
			
			colDens=wks_test_2.Columns(5);
			colA.SetLongName("par A");
			colB.SetLongName("par B");
			colC.SetLongName("par C");
			colD.SetLongName("Temperature " + pgOrigName);
			colD.SetUnits("eV");
			colDens.SetLongName("Density " + pgOrigName);
			colDens.SetUnits("sm^-3");
			break;
		case 1:
			if(FirstAnalyse_or_Second == 0) newPgName.Format("Сетка Т %s %s %s", str_SVCH_or_VCH, str_He_or_Ar, pgOrigName);
			if(FirstAnalyse_or_Second == 1) newPgName.Format("Сетка Ш %s %s %s", str_SVCH_or_VCH, str_He_or_Ar, pgOrigName);
			
			wks_test_3.Columns(1).SetLongName("Max Value");
			wks_test_3.Columns(1).SetUnits("Amper");
			wks_test_3.Columns(2).SetLongName("Temperature " + pgOrigName);
			wks_test_3.Columns(2).SetUnits("eV");
			wks_test_3.Columns(3).SetLongName("PeakVoltage " + pgOrigName);
			wks_test_3.Columns(3).SetUnits("Volts");
			wks_test_3.Columns(4).SetLongName("Energy " + pgOrigName);
			wks_test_3.Columns(4).SetUnits("eV");
			break;
		case 2:
			if(FirstAnalyse_or_Second == 0) newPgName.Format("Цилиндр %s %s %s", str_SVCH_or_VCH, str_He_or_Ar, pgOrigName);
			if(FirstAnalyse_or_Second == 1) newPgName.Format("Магнит %s %s %s", str_SVCH_or_VCH, str_He_or_Ar, pgOrigName);
			
			colA.SetLongName("Max Value");
			colA.SetUnits("Amper");
			colB.SetLongName("Temperature " + pgOrigName);
			colB.SetUnits("eV");
			colC.SetLongName("PeakVoltage " + pgOrigName);
			colC.SetUnits("Volts");
			colD.SetLongName("Energy " + pgOrigName);
			colD.SetUnits("eV");
			break;
	}
	pg.SetLongName(FindPageNameNumber(newPgName));
	
	if(Chast_to_Filtrate != 0 && z_s_c != 1)
	{
		wks_test_3.Columns(0).SetComments("0");
		wks_test_3.Columns(0).SetType(OKDATAOBJ_DESIGNATION_X);
		wks_test_3.Columns(0).SetUnits(wksUnits);
		strFilt.Format("%.2f", Chast_to_Filtrate*2);
	}
	
	Tree trFormat;  
	trFormat.Root.CommonStyle.Fill.FillColor.nVal = 18; 
	DataRange dr;
	otsecheniya.Format("S %.2f E %.2f", otsech_nachala, otsech_konca);
	for(int i_2=1, k=0; i_2<kolvo_otrezkov; i_2++) //otsech_nachala//otsech_konca
	{
		string vremya_diapazon, num_col;
		
		vremya_diapazon.Format("%.4f - %.4f", vec_Otr_Start_Indxs[i_2-1]/chastota_discret_zonda, vec_Otr_Start_Indxs[i_2-1]/chastota_discret_zonda + 1.0/chastota_pily - 0.0001);
		wksUnits.Format("%i - %i", (int)vec_Otr_Start_Indxs[i_2-1], (int)(vec_Otr_Start_Indxs[i_2-1] + chastota_discret_zonda/chastota_pily) - 1 );
		
		num_col.Format("%i", i_2);
		if(z_s_c == 1)
		{
			wks_test.Columns(i_2).SetUnits(wksUnits);
			wks_test_1.Columns(i_2).SetUnits(strFilt);
			wks_test_2.Columns(i_2).SetUnits(otsecheniya);
			wks_test.Columns(i_2).SetLongName(vremya_diapazon);
			wks_test_1.Columns(i_2).SetLongName(vremya_diapazon);
			wks_test_2.Columns(i_2).SetLongName(vremya_diapazon);
			wks_test.Columns(i_2).SetComments(num_col);
			wks_test_1.Columns(i_2).SetComments(num_col);
			wks_test_2.Columns(i_2).SetComments(num_col);
		}
		else
		{
			wks_test.Columns(i_2).SetUnits(wksUnits);
			wks_test_1.Columns(i_2).SetUnits(otsecheniya);
			wks_test.Columns(i_2).SetLongName(vremya_diapazon);
			wks_test_1.Columns(i_2).SetLongName(vremya_diapazon);
			wks_test.Columns(i_2).SetComments(num_col);
			wks_test_1.Columns(i_2).SetComments(num_col);
			if(Chast_to_Filtrate != 0)
			{
				wks_test_3.Columns(i_2).SetUnits(strFilt);
				wks_test_3.Columns(i_2).SetLongName(vremya_diapazon);
				wks_test_3.Columns(i_2).SetComments(num_col);
			}
		}
	
		if(i_2 % 2 != 0)
		{
			dr.Add("X", wks_test, 0, i_2, -1, i_2);
			if (0==dr.UpdateThemeIDs(trFormat.Root)) bool bRet = dr.ApplyFormat(trFormat, true, true);
			dr.Reset();
			dr.Add("X", wks_test_1, 0, i_2, -1, i_2);
			if (0==dr.UpdateThemeIDs(trFormat.Root)) bool bRet = dr.ApplyFormat(trFormat, true, true);
			dr.Reset();
			if(z_s_c != 1)
			{
				dr.Add("X", wks_test_2, i_2, 0, i_2, -1);
				if (0==dr.UpdateThemeIDs(trFormat.Root)) bool bRet = dr.ApplyFormat(trFormat, true, true);
				dr.Reset();
				if(Chast_to_Filtrate != 0)
				{
					dr.Add("X", wks_test_3, 0, i_2, -1, i_2);
					if (0==dr.UpdateThemeIDs(trFormat.Root)) bool bRet = dr.ApplyFormat(trFormat, true, true);
					dr.Reset();
				}
			}
			else
			{
				dr.Add("X", wks_test_2, 0, i_2, -1, i_2);
				if (0==dr.UpdateThemeIDs(trFormat.Root)) bool bRet = dr.ApplyFormat(trFormat, true, true);
				dr.Reset();
				dr.Add("X", wks_test_3, i_2, 0, i_2, -1);
				if (0==dr.UpdateThemeIDs(trFormat.Root)) bool bRet = dr.ApplyFormat(trFormat, true, true);
				dr.Reset();
			}
		}
	}

	return -1;
}

void rezult_graphs(Page pg, int z_s_c)
{
	string pgOrigName = pg.GetLongName();
	Worksheet wks_sheet = WksIdxByName(pg, "Parameters");
	
	if(z_s_c == 0)
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
		drTemp.Add(wks_sheet, 4, "Y");
		
		drDens.Add(wks_sheet, 0, "X");
		drDens.Add(wks_sheet, 5, "Y");
		
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
	}
	else
	{
		GraphPage gpTemp, gpPV, gpEnergy;
		gpTemp.Create("Origin");
		gpPV.Create("Origin");
		if(z_s_c == 2) gpEnergy.Create("Origin");
		GraphLayer glTemp = gpTemp.Layers(0), glPV = gpPV.Layers(0), glEnergy = gpEnergy.Layers(0);
		if(z_s_c == 1)
		{
			gpTemp.SetLongName ( FindPageNameNumber ( get_str_from_str_list(pgOrigName, " ", 0) + " " + get_str_from_str_list(pgOrigName, " ", 1) + " " + get_str_from_str_list(pgOrigName, " ", 4) + '(' + get_str_from_str_list(pgOrigName, " ", 5) + ')' + " Температура" ) );
			gpPV.SetLongName ( FindPageNameNumber ( get_str_from_str_list(pgOrigName, " ", 0) + " " + get_str_from_str_list(pgOrigName, " ", 1) + " " + get_str_from_str_list(pgOrigName, " ", 4) + '(' + get_str_from_str_list(pgOrigName, " ", 5) + ')' + " Потениал плазмы" ) );
		}
		else
		{
			gpTemp.SetLongName ( FindPageNameNumber ( get_str_from_str_list(pgOrigName, " ", 0) + " " + get_str_from_str_list(pgOrigName, " ", 3) + '(' + get_str_from_str_list(pgOrigName, " ", 4) + ')' + " Температура" ) );
			gpPV.SetLongName ( FindPageNameNumber ( get_str_from_str_list(pgOrigName, " ", 0) + " " + get_str_from_str_list(pgOrigName, " ", 3) + '(' + get_str_from_str_list(pgOrigName, " ", 4) + ')' + " Потениал плазмы" ) );
			gpEnergy.SetLongName ( FindPageNameNumber ( get_str_from_str_list(pgOrigName, " ", 0) + " " + get_str_from_str_list(pgOrigName, " ", 3) + '(' + get_str_from_str_list(pgOrigName, " ", 4) + ')' + " Энергия" ) );
		}
		Tree tr;
		tr = glTemp.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
		tr.Root.Legend.Font.Size.nVal = 12;
		if(0 == glTemp.UpdateThemeIDs(tr.Root) ) 
		{
			glTemp.ApplyFormat(tr, true, true);
			glPV.ApplyFormat(tr, true, true);
			if(z_s_c == 2) glEnergy.ApplyFormat(tr, true, true);
		}
	
		DataRange drTemp, drPV, drEnergy;
		DataPlot dpTemp, dpPV, dpEnergy;
		tr.Reset();
		
		drTemp.Add(wks_sheet, 0, "X");
		drTemp.Add(wks_sheet, 2, "Y");
		
		drPV.Add(wks_sheet, 0, "X");
		drPV.Add(wks_sheet, 3, "Y");
		
		if(z_s_c == 2) 
		{
			drEnergy.Add(wks_sheet, 0, "X");
			drEnergy.Add(wks_sheet, 4, "Y");
		}
		
		glTemp.AddPlot(drTemp, IDM_PLOT_LINESYMB);
		glPV.AddPlot(drPV, IDM_PLOT_LINESYMB);
		if(z_s_c == 2) glEnergy.AddPlot(drEnergy, IDM_PLOT_LINESYMB);
		
		dpTemp = glTemp.DataPlots(0);
		dpPV = glPV.DataPlots(0);
		if(z_s_c == 2) dpEnergy = glEnergy.DataPlots(0);
		
		tr = dpTemp.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
		tr.Root.Line.Width.dVal=1.1;
		tr.Root.Line.Color.nVal=0;
		tr.Root.Symbol.Size.nVal=5;;
		tr.Root.Symbol.EdgeColor.nVal=0;
		
		dpTemp.ApplyFormat(tr, true, true);
		dpPV.ApplyFormat(tr, true, true);
		if(z_s_c == 2) dpEnergy.ApplyFormat(tr, true, true);
		
		glTemp.Rescale();
		glPV.Rescale();
		if(z_s_c == 2) glEnergy.Rescale();
		
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
		//axisY.Scale.From.dVal = 0;
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
		//axisY.Scale.From.dVal = 0;
		axisY.Scale.IncrementBy.nVal = 1; // 0=increment by value; 1=number of major ticks
		axisY.Scale.MajorTicksCount.dVal = 10;
		
		if(z_s_c == 2) 
		{
			axisX = glEnergy.XAxis, axisY = glEnergy.YAxis;
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
			//axisY.Scale.From.dVal = 0;
			axisY.Scale.IncrementBy.nVal = 1; // 0=increment by value; 1=number of major ticks
			axisY.Scale.MajorTicksCount.dVal = 10;
		}
		
		glTemp.GraphObjects("XB").Text = "Время, с";
		glTemp.GraphObjects("YL").Text = "Температура ионов, эВ";
		glPV.GraphObjects("XB").Text = "Время, с";
		glPV.GraphObjects("YL").Text = "Потениал плазмы, В";
		if(z_s_c == 2) 
		{
			glEnergy.GraphObjects("XB").Text = "Время, с";
			glEnergy.GraphObjects("YL").Text = "Энергия, эВ";
		}
		
		set_active_layer(glTemp);
	}	
}

void create_labels(int z_s_c)
{
	GraphPage gp = Project.ActiveLayer().GetPage();
	GraphLayer gl1 = gp.Layers(0), gl2= gp.Layers(1), gl;
	if(gl2) gl = gl2;
	else gl = gl1;
	
	Worksheet wks_sheetLegend, wks1, wks2, wks_reflines;
	Get_OriginData_wks(wks1, wks2, wks_sheetLegend);
	
	wks_reflines = wks_sheetLegend.GetPage().Layers(1);
	
	Dataset slave, slave1;
	Curve crv;
	Tree tr;
	tr = gl.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
	
	if(z_s_c == 0)
	{
		if(wks_reflines.GetNumCols()<=4) wks_reflines.SetSize(-1,7);
		
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
	}
	else
	{
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
	}
    
	tr.Root.DataLabels.DataLabel1.Size.nVal = 12;
	tr.Root.DataLabels.DataLabel1.Color.nVal = 3;
	tr.Root.DataLabels.DataLabel1.YOffset.nVal = 40;
	tr.Root.DataLabels.DataLabel1.XOffset.nVal = 40;
	if (0==gl.UpdateThemeIDs(tr.Root)) gl.ApplyFormat(tr, true, true);
}

void create_cilin_graph_ogs()
{
	Worksheet wks = Project.ActiveLayer();
	
	vector vecPILA_C(wks.Columns(5)), vecPILA_M(wks.Columns(2)), vecTOK_C(wks.Columns(9)), vecTOK_M(wks.Columns(8));
	float otnosh_tok_pil, k;
	int i, n;
	
	if(vecPILA_C.GetSize() != vecTOK_C.GetSize())
	{
		otnosh_tok_pil = (float)vecTOK_C.GetSize()/vecPILA_C.GetSize();
		
		if(otnosh_tok_pil >= 2) /* откидываю все варианты пока разница размеров не будет хотябы в 2 раза, 
			так как счетчик заполнения типа int не будет работать с числом меньше 2 */
		{
			double chastota_discret_pily, chastota_discret_zonda;
			int Error = auto_chdsc(chastota_discret_pily, chastota_discret_zonda, wks.Columns(5), wks.Columns(9), ch_pil, -1, wks, 0);
			if( Error != -1 ) 
			{
				MessageBox(GetWindow(), FuncErrors(Error));
				return;
			}
			
			uint nDataSize = vecPILA_C.GetSize();
			vector vxPeaks(nDataSize), vyPeaks(nDataSize), vxData(nDataSize);
			vector<int> vnIndices(nDataSize);
			
			if(max(vecPILA_C) >= 0)
			{
				if( ocmath_find_peaks_by_local_maximum( &nDataSize, vxData, vecPILA_C, 
					vxPeaks, vyPeaks, vnIndices, POSITIVE_DIRECTION, (float)NumPointsOfOriginalPila*0.9) < OE_NOERROR ) return;
			}
			else
			{
				if( ocmath_find_peaks_by_local_maximum( &nDataSize, vxData, vecPILA_C, 
					vxPeaks, vyPeaks, vnIndices, NEGATIVE_DIRECTION, (float)NumPointsOfOriginalPila*0.9) < OE_NOERROR ) return;
			}
			if(nDataSize == 0) return;
			vnIndices.SetSize(nDataSize);
			
			int prav_naklon_ili_lev_naklon; // правый наклон означает что линия выглядит как \, а левый - /
			
			if(max(vecPILA_C) >= 0)
			{
				if( abs(vecPILA_C[vnIndices[1] + NumPointsOfOriginalPila*0.05]) / abs(vecPILA_C[vnIndices[1]]) > 0.5) prav_naklon_ili_lev_naklon = 0;
				else prav_naklon_ili_lev_naklon = 1;
			}
			else
			{
				if( abs(vecPILA_C[vnIndices[1] - NumPointsOfOriginalPila*0.05]) / abs(vecPILA_C[vnIndices[1]]) > 0.5) prav_naklon_ili_lev_naklon = 0;
				else prav_naklon_ili_lev_naklon = 1;
			}
			
			NumPointsOfOriginalPila = vnIndices[1] - vnIndices[0];
			
			/* сравниваем индексы по фазе с током */
			vnIndices *= otnosh_tok_pil;
			
			vector vsp;
			vsp = fit_linear_pila(vecPILA_C, vnIndices[0], vnIndices[1] - vnIndices[0]);
			
			/////////////////// первый член vecSupp - дельта между двумя соседними значениями пилы, второй - максимум, третий минимум ///////////////////
			vector vecSupp(3);
			vecSupp[0] = (max(vsp) - min(vsp))/(vsp.GetSize()-1);
			vecSupp[1] = max(vsp);
			vecSupp[2] = min(vsp);
			/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			
			vecPILA_C.SetSize(vecTOK_M.GetSize());
			if( prav_naklon_ili_lev_naklon == 0 )
			{
				/* заполняем основную часть */
				vecPILA_C[vnIndices[0]] = vecSupp[1];
				for(i = vnIndices[0]+1, n = 1; i < vecPILA_C.GetSize() && n < vnIndices.GetSize(); i++) 
				{
					if( i == vnIndices[n] )
					{
						vecPILA_C[i] = vecSupp[1];
						n++;
					}
					else vecPILA_C[i] = vecPILA_C[i-1] - vecSupp[0];
				}
				/* заполняем кусочек в начале */
				vecPILA_C[vnIndices[0]-1]= vecSupp[2];
				for(i = vnIndices[0]-2, n = 0; i > 0; i--, n++)
				{
					if( n == vnIndices[1] - vnIndices[0] )
					{
						vecPILA_C[i] = vecSupp[2];
						n = 0;
					}
					else vecPILA_C[i] = vecPILA_C[i+1] + vecSupp[0];
				}
				/* заполняем кусочек в конце */
				vecPILA_C[vnIndices[vnIndices.GetSize()-1]+1]= vecSupp[1];
				for(i = vnIndices[vnIndices.GetSize()-1]+2, n = 0; i < vecPILA_C.GetSize()-1; i++, n++)
				{
					if( n == vnIndices[1] - vnIndices[0] )
					{
						vecPILA_C[i] = vecSupp[1];
						n = 0;
					}
					else vecPILA_C[i] = vecPILA_C[i-1] - vecSupp[0];
				}
				vecPILA_C[vecPILA_C.GetSize()-1] = vecPILA_C[vecPILA_C.GetSize()-2];
			}
			else
			{
				/* заполняем основную часть */
				vecPILA_C[vnIndices[0]] = vecSupp[2];
				for(i = vnIndices[0]+1, n = 1; i < vecPILA_C.GetSize() && n < vnIndices.GetSize(); i++) 
				{
					if( i == vnIndices[n] )
					{
						vecPILA_C[i] = vecSupp[2];
						n++;
					}
					else vecPILA_C[i] = vecPILA_C[i-1] + vecSupp[0];
				}
				/* заполняем кусочек в начале */
				vecPILA_C[vnIndices[0]-1]= vecSupp[1];
				for(i = vnIndices[0]-2, n = 0; i > 0; i--, n++)
				{
					if( n == vnIndices[1] - vnIndices[0] )
					{
						vecPILA_C[i] = vecSupp[1];
						n = 0;
					}
					else vecPILA_C[i] = vecPILA_C[i+1] - vecSupp[0];
				}
				/* заполняем кусочек в конце */
				vecPILA_C[vnIndices[vnIndices.GetSize()-1]+1]= vecSupp[2];
				for(i = vnIndices[vnIndices.GetSize()-1]+2, n = 0; i < vecPILA_C.GetSize()-1; i++, n++)
				{
					if( n == vnIndices[1] - vnIndices[0] )
					{
						vecPILA_C[i] = vecSupp[2];
						n = 0;
					}
					else vecPILA_C[i] = vecPILA_C[i-1] + vecSupp[0];
				}
				vecPILA_C[vecPILA_C.GetSize()-1] = vecPILA_C[vecPILA_C.GetSize()-2];
			}
		}
		if(otnosh_tok_pil <= 0.5) /* откидываю все варианты пока разница размеров не будет хотябы в 2 раза, 
			так как счетчик заполнения типа int не будет работать с числом меньше 2 */
		{
			for(i = 0 + (int)(1/otnosh_tok_pil); i < vecPILA_C.GetSize(); i += (int)(1/otnosh_tok_pil)) vecPILA_C.RemoveAt(i);	
		}
	}
	
	if(vecPILA_M.GetSize() != vecTOK_M.GetSize()) 
	{
		otnosh_tok_pil = (float)vecTOK_M.GetSize()/vecPILA_M.GetSize();
		
		if(otnosh_tok_pil >= 2) /* откидываю все варианты пока разница размеров не будет хотябы в 2 раза, 
			так как счетчик заполнения типа int не будет работать с числом меньше 2 */
		{
			double chastota_discret_pily, chastota_discret_zonda;
			int Error = auto_chdsc(chastota_discret_pily, chastota_discret_zonda, wks.Columns(2), wks.Columns(8), ch_pil, -1, wks, 1);
			if( Error != -1 ) 
			{
				MessageBox(GetWindow(), FuncErrors(Error));
				return;
			}
			
			uint nDataSize = vecPILA_M.GetSize();
			vector vxPeaks(nDataSize), vyPeaks(nDataSize), vxData(nDataSize);
			vector<int> vnIndices(nDataSize);
			
			if(max(vecPILA_M) >= 0)
			{
				if( ocmath_find_peaks_by_local_maximum( &nDataSize, vxData, vecPILA_M, 
					vxPeaks, vyPeaks, vnIndices, POSITIVE_DIRECTION, (float)NumPointsOfOriginalPila*0.9) < OE_NOERROR ) return;
			}
			else
			{
				if( ocmath_find_peaks_by_local_maximum( &nDataSize, vxData, vecPILA_M, 
					vxPeaks, vyPeaks, vnIndices, NEGATIVE_DIRECTION, (float)NumPointsOfOriginalPila*0.9) < OE_NOERROR ) return;
			}
			if(nDataSize == 0) return;
			vnIndices.SetSize(nDataSize);
			
			int prav_naklon_ili_lev_naklon; // правый наклон означает что линия выглядит как \, а левый - /
			
			if(max(vecPILA_M) >= 0)
			{
				if( abs(vecPILA_M[vnIndices[1] + NumPointsOfOriginalPila*0.05]) / abs(vecPILA_M[vnIndices[1]]) > 0.5) prav_naklon_ili_lev_naklon = 0;
				else prav_naklon_ili_lev_naklon = 1;
			}
			else
			{
				if( abs(vecPILA_M[vnIndices[1] - NumPointsOfOriginalPila*0.05]) / abs(vecPILA_M[vnIndices[1]]) > 0.5) prav_naklon_ili_lev_naklon = 0;
				else prav_naklon_ili_lev_naklon = 1;
			}
			
			NumPointsOfOriginalPila = vnIndices[1] - vnIndices[0];
			otsech_nachala = otsech_konca = 0.3;
			
			vector vsp;
			vsp = fit_linear_pila(vecPILA_M, vnIndices[0], (vnIndices[1] - vnIndices[0])*otnosh_tok_pil);
			
			/* сравниваем индексы по фазе с током */
			vnIndices *= otnosh_tok_pil;
			
			/////////////////// первый член vecSupp - дельта между двумя соседними значениями пилы, второй - максимум, третий минимум ///////////////////
			vector vecSupp(3);
			float a = min(vsp), b = max(vsp);
			vecSupp[1] = vsp[0] - ( (vsp[vsp.GetSize()-1] - vsp[0]) / (1-otsech_nachala-otsech_konca) ) * otsech_nachala; // Начало
			vecSupp[2] = vsp[vsp.GetSize()-1] + ( (vsp[vsp.GetSize()-1] - vsp[0]) / (1-otsech_nachala-otsech_konca) ) * otsech_nachala; // Конец
			vecSupp[0] = (vecSupp[2] - vecSupp[1])/(vsp.GetSize()/(1-otsech_nachala-otsech_konca) - 1);
			/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			
			vecPILA_M.SetSize(vecTOK_M.GetSize());
			/* заполняем основную часть */
			vecPILA_M[vnIndices[0]] = vecSupp[1];
			for(i = vnIndices[0]+1, n = 1; i < vecPILA_M.GetSize() && n < vnIndices.GetSize(); i++) 
			{
				if( i == vnIndices[n] )
				{
					vecPILA_M[i] = vecSupp[1];
					n++;
				}
				else vecPILA_M[i] = vecPILA_M[i-1] + vecSupp[0];
			}
			/* заполняем кусочек в начале */
			vecPILA_M[vnIndices[0]-1]= vecSupp[2];
			for(i = vnIndices[0]-2, n = 0; i > 0; i--, n++)
			{
				if( n == vnIndices[1] - vnIndices[0] )
				{
					vecPILA_M[i] = vecSupp[2];
					n = 0;
				}
				else vecPILA_M[i] = vecPILA_M[i+1] - vecSupp[0];
			}
			/* заполняем кусочек в конце */
			vecPILA_M[vnIndices[vnIndices.GetSize()-1]+1]= vecSupp[1];
			for(i = vnIndices[vnIndices.GetSize()-1]+2, n = 0; i < vecPILA_M.GetSize()-1; i++, n++)
			{
				if( n == vnIndices[1] - vnIndices[0] )
				{
					vecPILA_M[i] = vecSupp[1];
					n = 0;
				}
				else vecPILA_M[i] = vecPILA_M[i-1] + vecSupp[0];
			}
			vecPILA_M[vecPILA_M.GetSize()-1] = vecPILA_M[vecPILA_M.GetSize()-2];
			
			//if( prav_naklon_ili_lev_naklon == 0 )
			//{
				///* заполняем основную часть */
				//vecPILA_M[vnIndices[0]] = vecSupp[1];
				//for(i = vnIndices[0]+1, n = 1; i < vecPILA_M.GetSize() && n < vnIndices.GetSize(); i++) 
				//{
					//if( i == vnIndices[n] )
					//{
						//vecPILA_M[i] = vecSupp[1];
						//n++;
					//}
					//else vecPILA_M[i] = vecPILA_M[i-1] - vecSupp[0];
				//}
				///* заполняем кусочек в начале */
				//vecPILA_M[vnIndices[0]-1]= vecSupp[2];
				//for(i = vnIndices[0]-2, n = 0; i > 0; i--, n++)
				//{
					//if( n == vnIndices[1] - vnIndices[0] )
					//{
						//vecPILA_M[i] = vecSupp[2];
						//n = 0;
					//}
					//else vecPILA_M[i] = vecPILA_M[i+1] + vecSupp[0];
				//}
				///* заполняем кусочек в конце */
				//vecPILA_M[vnIndices[vnIndices.GetSize()-1]+1]= vecSupp[1];
				//for(i = vnIndices[vnIndices.GetSize()-1]+2, n = 0; i < vecPILA_M.GetSize()-1; i++, n++)
				//{
					//if( n == vnIndices[1] - vnIndices[0] )
					//{
						//vecPILA_M[i] = vecSupp[1];
						//n = 0;
					//}
					//else vecPILA_M[i] = vecPILA_M[i-1] - vecSupp[0];
				//}
				//vecPILA_M[vecPILA_M.GetSize()-1] = vecPILA_M[vecPILA_M.GetSize()-2];
			//}
			//else
			//{
				///* заполняем основную часть */
				//vecPILA_M[vnIndices[0]] = vecSupp[2];
				//for(i = vnIndices[0]+1, n = 1; i < vecPILA_M.GetSize() && n < vnIndices.GetSize(); i++) 
				//{
					//if( i == vnIndices[n] )
					//{
						//vecPILA_M[i] = vecSupp[2];
						//n++;
					//}
					//else vecPILA_M[i] = vecPILA_M[i-1] + vecSupp[0];
				//}
				///* заполняем кусочек в начале */
				//vecPILA_M[vnIndices[0]-1]= vecSupp[1];
				//for(i = vnIndices[0]-2, n = 0; i > 0; i--, n++)
				//{
					//if( n == vnIndices[1] - vnIndices[0] )
					//{
						//vecPILA_M[i] = vecSupp[1];
						//n = 0;
					//}
					//else vecPILA_M[i] = vecPILA_M[i+1] - vecSupp[0];
				//}
				///* заполняем кусочек в конце */
				//vecPILA_M[vnIndices[vnIndices.GetSize()-1]+1]= vecSupp[2];
				//for(i = vnIndices[vnIndices.GetSize()-1]+2, n = 0; i < vecPILA_M.GetSize()-1; i++, n++)
				//{
					//if( n == vnIndices[1] - vnIndices[0] )
					//{
						//vecPILA_M[i] = vecSupp[2];
						//n = 0;
					//}
					//else vecPILA_M[i] = vecPILA_M[i-1] + vecSupp[0];
				//}
				//vecPILA_M[vecPILA_M.GetSize()-1] = vecPILA_M[vecPILA_M.GetSize()-2];
			//}	
		}
		if(otnosh_tok_pil <= 0.5) /* откидываю все варианты пока разница размеров не будет хотябы в 2 раза, 
			так как счетчик заполнения типа int не будет работать с числом меньше 2 */
		{
			for(i = 0 + (int)(1/otnosh_tok_pil); i < vecPILA_M.GetSize(); i += (int)(1/otnosh_tok_pil)) vecPILA_M.RemoveAt(i);	
		}
	}
	
	Dataset<double> dsPILA_C_d(wks.Columns(5)), dsPILA_M_d(wks.Columns(2)); 
	Dataset<float> dsPILA_C_f(wks.Columns(5)), dsPILA_M_f(wks.Columns(2));
	
	if(wks.Columns(5).GetInternalData() == FSI_DOUBLE) 
		dsPILA_C_d = vecPILA_C;
	else
		dsPILA_C_f = vecPILA_C;
	
	if(wks.Columns(2).GetInternalData() == FSI_DOUBLE) 
		dsPILA_M_d = vecPILA_M;
	else
		dsPILA_M_f = vecPILA_M;
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