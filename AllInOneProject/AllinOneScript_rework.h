#include <Origin.h>
#include <GetNBox.h>
#include <..\Originlab\graph_utils.h> // Needed for page_add_layer function

#define NOGAUSSFITS // будет ли использована аппроксимация гауссом в обработке сеточного
#define NOGAUSSFITC // будет ли использована аппроксимация гауссом в обработке цилиндра



/* INNER FUNCTIONS */
int data_process(int, string, Worksheet &, Column &, Column &);	// MathFuncs
int linfit(vector &, vector &, double, int, vector &);			// MathFuncs
bool is_signalpeakslookingdown(vector &);						// MathFuncs

/**********************************************************************************/

string ERR_GetErrorDescription(int err);						// SubFuncs
Page & create_book(int,
				   string,
				   int,
				   int,
				   vector &,
				   vector<int> &,
				   string,
				   string,
				   string);										// SubFuncs
string FindPageNameNumber(string);								// SubFuncs
Worksheet & WksIdxByName(Page &, string);						// SubFuncs
Column & ColIdxByName(Worksheet &, string);						// SubFuncs
Column & ColIdxByUnit(Worksheet &, string);						// SubFuncs
int Get_OriginData_wks(Page &, Page &, Page &);					// SubFuncs
void compare_books(uint, vector <string> &);					// SubFuncs

/**********************************************************************************/

int swap_graph(int, int);										// Swap_Graph

/**********************************************************************************/

void rezult_graphs(Page &, int);								// SubFuncsForGraph
int set_color_rnd();											// SubFuncsForGraph	
int set_color_except_bad(int);									// SubFuncsForGraph
void create_labels(int);										// SubFuncsForGraph
int take_selection(vector <int> &, int &);						// SubFuncsForGraph
GraphPage create_empty_graph(Page &,
							 vector <int> &);					// SubFuncsForGraph
int fill_single_segment_graph(GraphPage &,
							  Page &,
							  vector <int> &);					// SubFuncsForGraph
int fill_compare_graph(GraphPage &,
					   vector <int> &,
					   Page &);									// SubFuncsForGraph
int fill_evolution_graph(GraphPage &,
					     Page &);								// SubFuncsForGraph
Page & create_legend(Page &);									// SubFuncsForGraph
int fill_legend(Page &,
				Page &,
				vector <int> &);								// SubFuncsForGraph
	
/**********************************************************************************/

int interf_prog(Column &, int, Worksheet &, double, int);		// Interf_process

/**********************************************************************************/

int graphZond();												// Zond_Graph
int make_add_data_for_graph_zond(Page &,
								 Page &,
								 vector <int> &);				// Zond_Graph
int add_sublines_to_graph_zond(GraphPage &,
							   Page &);							// Zond_Graph
int reprocess_zond();											// Zond_Graph

/**********************************************************************************/

int graphDoubleProbe();											// DoubleProbe_Graph
int make_add_data_for_graph_doubleprobe(Page &,
								 Page &,
								 vector <int> &);				// DoubleProbe_Graph
int add_sublines_to_graph_doubleprobe(GraphPage &,
									  Page &);					// DoubleProbe_Graph
int reprocess_doubleprobe();									// DoubleProbe_Graph

/**********************************************************************************/

int graphSetoch();												// Setka_Graph
int make_add_data_for_graph_setka(Page &,
								  Page &,
								  vector <int> &);				// Setka_Graph
int add_sublines_to_graph_setka(GraphPage &,
							    Page &);						// Setka_Graph
int reprocess_setka();											// Setka_Graph

/**********************************************************************************/

int graphCilin();												// Cilin_Graph
int make_add_data_for_graph_cilin(Page &,
								  Page &,
								  vector <int> &);				// Cilin_Graph
int add_sublines_to_graph_cilin(GraphPage &,
							    Page &);						// Cilin_Graph
int reprocess_cilin();											// Cilin_Graph

/**********************************************************************************/

int dlg_ColumnNotFound(Worksheet &,
					   string);									// Dialogs
int dlg_ZondProcess(vector <int> &,
					vector <int> &,
					vector <int> &,
					vector &,
					vector <string> &,
					vector <string> &,
					vector <int> &,
					int &);										// Dialogs
int dlg_SetkaProcess(vector <int> &,
					 vector <int> &,
					 vector <int> &,
					 vector <string> &,
					 vector <string> &,
					 vector <int> &,
					 int &);									// Dialogs
int dlg_CilinProcess(vector <int> &,
					 vector <int> &,
					 vector <int> &,
					 vector <string> &,
					 vector <string> &,
					 vector <int> &,
					 int &);									// Dialogs
int dlg_DoubleProbeProcess(vector <int> &,
					       vector <int> &,
					       vector <int> &,
					       vector &,
					       vector <string> &,
					       vector <string> &,
					       vector <int> &,
						   int &);								// Dialogs

/**********************************************************************************/



/* GLOBAL VARIABLES */
extern int ch_dsc_p, 
	       ch_dsc_t, 
	       freqP, 
	       r_z1, 
	       r_z2, 
	       r_s1, 
	       r_s2, 
	       r_cil, 
	       Num_iter, 
	       coefPila, 
	       ch_dsc_ntrfrmtr, 
	       ImpulseType,
	       FuelType,
	       resistance,
	       Number_of_rezult_plasma_parametrs;
extern double leftP, 
			  rightP, 
			  S1, 
			  S2, 
			  M_Ar, 
			  M_He,
			  razbros_po_pile, 
			  linfitP, 
			  st_time_end_time[2],
			  filtS =,
			  S;	
extern string s_pz, 
			  s_tz, 
			  s_tz2, 
			  s_ps, 
			  s_tst, 
			  s_tssh,
			  svch, 
			  vch, 
			  icr, 
			  s_pc, 
			  s_tc, 
			  s_ntrfrmtr,
			  s_pm, 
			  s_tm,
			  s_currp;
 



/* Error codes */
enum 
{
    ERR_NoError = 0,                // No error
    ERR_BadInputVecs = -6201,       // Corrupted input vectors
    ERR_ZeroInputVals = -6202,      // Input data error, some values equals 0
    ERR_BadCutOff = -6203,          // Сut-off points must be less than 0.9 in total and more than 0 each
    ERR_BadLinFit = -6204,          // The length of linear fit must be less than 0.9
    ERR_BadFactorizing = -6205,     // Error after Pila|Signal factorizing
    ERR_BadNoise = -6206,           // Error after noise extracting
    ERR_BadSegInput = -6207,        // Input segment's values error while segmend approximating
    ERR_TooFewSegs = -6208,         // Less then 4 segments found, check input arrays  
    ERR_BadSegsLength = -6209,      // Error in finding segments length, check input params
    ERR_BadLinearPila = -6210,      // Error in pila linearizing, check cut-off params
    ERR_TooManyAttempts = -6211,    // More than 5 attempts to find signal, check if signal is noise or not
    ERR_BadStartEnd = -6212,        // Error in finding start|end of signal, check if signal is noise or not
    ERR_TooManySegs = -6213,        // Too many segments, check if signal is noise or not
    ERR_NoSegs = -6214,             // No segments found, check if signal is noise or not
    ERR_BadDiagNum = -6215,         // Diagnostics number must be > 0 and < 2
    ERR_IdxOutOfRange = -6216,      // Index is out of range. Programm continued running
    ERR_Exception = -6217,          // An exception been occured whule running, script did not stopt working
    ERR_BadStEndTime = -6218,       // Error!: start time must be less then end time and total time, more than 0\n\
                                    end time must be less then total time, more then 0
	ERR_BufferExtension = -6219     // Buffer extension, bad impulse
};




//#pragma dll(PlasmaDataProcessing, header)
#pragma dll(PlasmaDataProcessing, header)
int OriginFindSignal(int, double*, double*, double*, uint, int*, int*, double*, int*);
int OriginAll(int, double*, double*, double*, double*, double*, double*, double*, double*);
int OriginMakeOne(int, double*, double*, double*, double*, double*, double*, double*);