/*! \file: MGAFeedAssayProcessor.cpp
\author Zitong Wang
\date 2020-09-19
\version 1.0.0
\note: Developed for automating MGA Feed Assay result processing
\compiled with std:c++17
\This program takes in the exported standard and unknown results from Empower, sets up calibration curve and calculate assay results/recoveries.
*/

#include <iostream>
#include <fstream>
#include <string>
#include <vector>
#include <locale>
#include <filesystem>
#include <map>
#include <iterator>
#include <regex>
#include <iomanip>
#include <sstream>
#include <math.h>
#include <ctime>
#include "libxl.h"

#pragma warning(disable : 4996)
using namespace std;
using namespace std::filesystem;
using namespace libxl;


const string VERSION = "1.0.1";
int sampleTypeIndex = -1;
int sampleNameIndex = -1;
int peakNameIndex = -1;
int RTIndex = -1;
int areaIndex = -1;
int weightIndex = -1;
int dilutionIndex = -1;
double expectedPotency = 2000.0;
size_t standardInjectionsCount = 0;
double standardPeakRatioMean = 0.0;
double standardStev = 0.0;
double standardRSD = 0.0;
bool calculateRecovery = false;
double standardConc = 1.0;
bool averageMode = false;


struct injection {
	string sampleType;
	double melengestrolArea = 0;
	double megestrolArea = 0;
	double peakRatio = 0.0;
	double dilution = 1.0;
	double weight = 1.0;
	double assay = 0.0;
	double recovery = 0.0;
};


void help() {
	cout << endl << "Auto Processor and report generator for MGA Feed Assay" << endl << endl
		<< "Usage: MGAFeedAssayProcessor [--help|version] infile" << endl << endl
		<< "Examples:" << endl
		<< "  MGAFeedAssayProcessor myfile.txt" << endl
		<< "    > reads data from myfile.txt exported from Empower and perfroms calculation " << endl << endl
		<< "  MGAFeedAssayProcessor --help" << endl
		<< "    > displays the help" << endl << endl
		<< "  rle --version" << endl
		<< "    > displays version number in the format d.d.d" << endl
		<< endl << "Note: this program uses the third-party library libxl to generate output excel files." << endl
		<< "Please make sure libxl.dll is placed under the same directory as the program" << endl
		<< "Excel file will be generated under the same directory with name format inputFileName + CurrentDate(YYYYMD) + xls." << endl
		<< "File with same name under the directory "
		<< "will be overwritten." << endl << endl
		<< "source code of the program and most updated version can be found under" << endl
		<< "https://github.com/zwang452/MGAFeedAssayProcessor" << endl;
}


void split(const string& s, char delim, map<string, injection>& results) {
	stringstream ss;
	ss.str(s);
	string item;
	injection injectionData;  
	string sampleName;
	vector<std::string> elems;
	int indexCount = -1;
	bool isDataRow = false;
	while (getline(ss, item, delim)) {
		elems.push_back(item);
		indexCount++;

		if (item.find("Sample Type") != string::npos) {
			sampleTypeIndex = indexCount;
		}
		else if (item.find("Name") != string::npos && item.find("SampleName") == string::npos) {
			peakNameIndex = indexCount;
		}
		else if (item.find("RT") != string::npos) {
			RTIndex = indexCount;
		}
		else if (item.find("Area") != string::npos) {
			areaIndex = indexCount;
		}
		else if (item.find("SampleName") != string::npos) {
			sampleNameIndex = indexCount;
		}
		else if (item.find("SampleWeight") != string::npos) {
			weightIndex = indexCount;
		}
		else if (item.find("Dilution") != string::npos) {
			dilutionIndex = indexCount;
		}
		else {
			isDataRow = true;
		}

		if (item.find("Peak Results") != string::npos) {
			isDataRow = false;
		}
	}
	if (isDataRow) {
		sampleName = elems.at(sampleNameIndex);
		auto loc = results.find(sampleName);
		if (loc == results.end()) {
			injectionData.sampleType = elems.at(sampleTypeIndex);
			if (injectionData.sampleType.find("tandard") != string::npos) {
				standardInjectionsCount++;
			}
			if (weightIndex != -1) {
				injectionData.weight = stod(elems.at(weightIndex), nullptr);
			}
			if (dilutionIndex != -1) {
				injectionData.dilution = stod(elems.at(dilutionIndex), nullptr);
			}
			if (elems.at(peakNameIndex).find("Melengestrol") != string::npos) {
				injectionData.melengestrolArea = stod(elems.at(areaIndex), nullptr);
			}
			else if (elems.at(peakNameIndex).find("Megestrol") != string::npos) {
				injectionData.megestrolArea = stod(elems.at(areaIndex), nullptr);
			}
			results.insert(pair<string, injection>(sampleName, injectionData));
		}
		else {
			if (elems.at(peakNameIndex).find("Melengestrol") != string::npos) {
				(*loc).second.melengestrolArea = stod(elems.at(areaIndex), nullptr);
			}
			else if (elems.at(peakNameIndex).find("Megestrol") != string::npos) {
				(*loc).second.megestrolArea = stod(elems.at(areaIndex), nullptr);
			}
		}

	}

}

void mannually_enter_weights(map<string, injection>& results) {
	for (auto &result : results) {
		if (result.second.sampleType.find("Unknown") != string::npos) {
			std::cout << endl << "Please enter sample weight in grams:" << endl <<
				"sample: " << result.first << endl;
			while (!(cin >> result.second.weight)) {
				cerr << "Invalid input. Try again: ";
				cin.clear();
				cin.ignore(numeric_limits<streamsize>::max(), '\n');
			}
		}
	}
}

void manually_enter_dilutions(map<string, injection>& results) {
	for (auto &result : results) {
		if (result.second.sampleType.find("Unknown") != string::npos) {
			std::cout << endl << "Please enter dilution factor for sample "
				<< result.first << endl;
			while (!(cin >> result.second.dilution)) {
				cerr << "Invalid input. Try again: ";
				cin.clear();
				cin.ignore(numeric_limits<streamsize> ::max(), '\n');
			}

		}
	}
}

void calculate_calibration_curve(map<string, injection>&results) {
	double standardPeakRatioSum = 0.0;
	for (auto &result : results) {
		if (result.second.sampleType.find("tandard") != string::npos) {
			standardPeakRatioSum += result.second.peakRatio;
		}
	}

	standardPeakRatioMean = standardPeakRatioSum / standardInjectionsCount;

	for (auto result : results) {
		if (result.second.sampleType.find("tandard") != string::npos) {
			standardStev += pow((result.second.peakRatio - standardPeakRatioMean),2);
		}
	}
	standardStev /= (standardInjectionsCount-1);
	standardStev = sqrt(standardStev);
	standardRSD = standardStev / standardPeakRatioMean;
}

void calculate_assay(map<string, injection>& results) {
	for (auto& result : results) {
		if (result.second.sampleType.find("Unknown") != string::npos) {
			result.second.assay = (result.second.peakRatio / standardPeakRatioMean) * standardConc * (result.second.dilution / result.second.weight);
			if (calculateRecovery == true) {
				result.second.recovery = result.second.assay / expectedPotency;
			}	
		}
	}
}

void export_to_xml(map<string, injection>& results, string& inputFile) {

	time_t now = time(0);
	char buffer[26];
	char* stringNow = ctime(&now);
	wchar_t wcharNow[100];
	mbstowcs(wcharNow, stringNow, 100);
	

	//Format date
	tm* ltm = localtime(&now);
	int month = 1+ ltm->tm_mon;
	int day = ltm->tm_mday;
	int year = 1900 + ltm->tm_year;
	string outputNameString = inputFile.substr(0, inputFile.length() - 3) + to_string(year) + to_string(month) + to_string(day) + ".xls  " ;
	wstring outputNameWstring = wstring(outputNameString.begin(), outputNameString.end());
	const wchar_t* outputWchar = outputNameWstring.c_str();
	//Create new instance of xml file
	Book* book = xlCreateBook();

	Format* percentage = book->addFormat();
	percentage->setNumFormat(NUMFORMAT_PERCENT_D2);

	Font* redFont = book->addFont();
	redFont->setColor(COLOR_RED);
	Format* redFontFormat = book->addFormat();
	redFontFormat -> setFont(redFont);
	redFontFormat ->setBorderColor(COLOR_RED);

	Font* greenFont = book->addFont();
	greenFont->setColor(COLOR_GREEN);
	Format* greenFontFormat = book->addFormat();
	greenFontFormat->setFont(greenFont);
	greenFontFormat->setBorderColor(COLOR_GREEN);


	if (book) {
		Sheet* sheet = book->addSheet(L"Sheet1");
		if (sheet) {
			int col = 0;
			int row = 1;
			sheet->writeStr(row++, col, L"Report auto-generated by MGA Feed Assay Processor, Version 1.0.1, Eric Wang, BAM ");
			sheet->writeStr(row++, col, wcharNow);
			sheet->writeStr(row, col++, L"Standard No.");
			sheet->writeStr(row, col++, L"MGA Peak Area");
			sheet->writeStr(row, col++, L"MEG Peak Area");
			sheet->writeStr(row++, col, L"Peak Ratio");
			col = 0;

			size_t stdNum = 0;
			for (auto result : results) {
				if (result.second.sampleType.find("tandard") != string::npos) {
					stdNum++;
					sheet->writeNum(row, col++, stdNum);
					sheet->writeNum(row, col++, result.second.melengestrolArea);
					sheet->writeNum(row, col++, result.second.megestrolArea);
					sheet->writeNum(row++, col, result.second.peakRatio);
					col = 0;
				}
			}

			sheet->writeStr(row, col++, L"Peak Ratio Mean");
			sheet->writeNum(row++, col, standardPeakRatioMean);
			col = 0;

			sheet->writeStr(row, col++, L"Peak Ratio Stdev");
			sheet->writeNum(row++, col, standardStev);
			col = 0;

			sheet->writeStr(row, col++, L"Peak Ratio RSD");
			sheet->writeNum(row, col++, standardRSD, percentage);
			if ((int)(standardRSD*100) <= 10) {
				sheet->writeStr(row++, col, L"RSD <=10% Passed!", greenFontFormat);
			}
			else {
				sheet->writeStr(row++, col, L"RSD >10% Failed!", redFontFormat);
			}
			
			col = 0;

			sheet->writeStr(row, col++, L"Std conc. (ppb)");
			sheet->writeNum(row++, col, standardConc);
			col = 0;


			sheet->writeStr(row++, col, L"Sample results");
			sheet->writeStr(row, col++, L"Sample name");
			sheet->writeStr(row, col++, L"MGA Peak Area");
			sheet->writeStr(row, col++, L"MEG Peak Area");
			sheet->writeStr(row, col++, L"Peak Ratio");
			sheet->writeStr(row, col++, L"Weight (g)");
			sheet->writeStr(row, col++, L"Dilution");
			sheet->writeStr(row, col++, L"Assay (ppb)");
			if (calculateRecovery) {
				sheet->writeStr(row, col++, L"Recovery");
			}
			
			if (averageMode == true) {
				sheet->writeStr(row, col++, L"Average Assay for Duplicates (ppb)");
				if (calculateRecovery) {
					sheet->writeStr(row, col++, L"Average Recovery for Duplicates");
				}
			}
			row++;
			col = 0;
			
			bool isSecondInjection = false;
			double firstAssay = 0.0;
			double firstRecovery = 0.0;
			
			
			for (auto result : results) {
				if (result.second.sampleType.find("Unknown") != string::npos) {
					
					sheet->writeStr(row, col++, wstring(result.first.begin(), result.first.end()).c_str());
					sheet->writeNum(row, col++, result.second.melengestrolArea);
					sheet->writeNum(row, col++, result.second.megestrolArea);
					sheet->writeNum(row, col++, result.second.peakRatio);
					sheet->writeNum(row, col++, result.second.weight);
					sheet->writeNum(row, col++, result.second.dilution);
					sheet->writeNum(row, col++, result.second.assay);
					
					if (calculateRecovery) {
						sheet->writeNum(row, col++, result.second.recovery, percentage);
					}

					if (averageMode == true && isSecondInjection) {
						sheet->writeNum(row, col++, (firstAssay + result.second.assay)/2.0);
						if (calculateRecovery) {
							sheet->writeNum(row, col++, (firstRecovery + result.second.recovery) / 2.0, percentage);
						}
						
						isSecondInjection = false;
					}
					else {
						isSecondInjection = true;
						firstAssay = result.second.assay;
						firstRecovery = result.second.recovery;
					}
					row++;
					col = 0;
				}
			}
			sheet->setCol(0, 0, 35);
			sheet->setCol(1, 7, 18);
			sheet->setCol(8, 9, 35);

			if (book->save(outputWchar))
			{
				std::cout << "Data output to excel file named " + outputNameString + " successfully!" << endl;
			}
			else
			{
				std::cout << book->errorMessage() << std::endl
					<< "Please check there is no file named " + outputNameString + "open under the same directory" << endl
					<< "No excel file generated" <<endl;
			}
			book->release();
		}
	}
}

int main(int argc, char* argv[]) {
	cout << " MGAFeedAssayProcessor (c) 2020, Eric Wang" << endl
		<< "==============================================================" << endl
		<< "Version " << VERSION << endl
		<< "zwang452@gmail.com" << endl;

	if (argc == 1) {
		cout << "Error: no data file entered - showing help." << endl;
		help();
		exit(EXIT_FAILURE);
	}
	string inputFile = argv[1];
	if (inputFile == "--help") {
		help();
		exit(EXIT_FAILURE);
	}
	else if (inputFile == "--version") {
		cout << "Version " << VERSION << endl;
		exit(EXIT_FAILURE);
	}

	map<string, injection> results;
	
	ifstream infile(inputFile);
	if (!infile) {
		cout << "Error: file <" << inputFile << "> doesn't exist" << endl;
		exit(EXIT_FAILURE);
	}
	string line;
	
	while (getline(infile, line, '\r')) {
		split(line, '\t', results);
	}
	cout << "Data parsing complete..." << endl;

	for (auto &result : results) {
		result.second.peakRatio =  result.second.melengestrolArea / result.second.megestrolArea;
	}
	cout << "Peak ratio calculation complete..." << endl;

	cout << "Previewing Data..." << endl;
	for (auto result : results) {
		//result.second.peakRatio = result.second.melengestrolArea / result.second.megestrolArea;
		cout << result.first << " " << result.second.sampleType << " " << result.second.megestrolArea << " " << result.second.melengestrolArea << " " <<
			result.second.peakRatio << endl;
	}

	if (standardInjectionsCount == 0) {
		cerr << "Missing standard Injection! Please re-export data files including standard injections";
		return(EXIT_FAILURE);
	}
	else {
		calculate_calibration_curve(results);
		cout << endl << "Calibration curve constructed..." << endl
			<< "Number of standard injections included: " << standardInjectionsCount << endl
			<< "Mean Peak Ratio: " << standardPeakRatioMean << endl
			<< "Standard Deviation: " << standardStev << endl
			<< "RSD: " << standardRSD * 100 << "%" << endl;
	}


	
	if (weightIndex == -1) {
		cout << endl <<"No weights info detected in exported files. Would you want to: " << endl
			<< "A. mannually enter them here" << endl
			<< "B. calculate using a default value of 1.0 g" << endl
			<< "press any other key to exit" <<endl;
		char selection;
		cin >> selection;
		cout << selection;
		if (tolower(selection) == 'a') {
			mannually_enter_weights(results);
		}
		else if (tolower(selection) != 'b') {
			return(EXIT_SUCCESS);
		}
	}

	if (dilutionIndex == -1) {
		cout << endl << "No dilution info detected in exported files. Would you want to: " << endl
			<< "A. mannually enter them here" << endl  
			<< "B. calculate using a default value of 1" << endl
			<< "C. calculate using a customered default value" << endl
			<< "press any other key to exit" << endl;
		char selection;
		size_t defaultDilution = 1;
		cin >> selection;
		cout << selection;
		if (tolower(selection) == 'a') {
			manually_enter_dilutions(results);
		}
		else if (tolower(selection) == 'c') {
			while (!(cin >> defaultDilution)) {
				cerr << "Invalid input. Try again: ";
				cin.clear();
				cin.ignore(numeric_limits<streamsize> ::max(), '\n');
			}
		}
		else if (tolower(selection) != 'b') {
			return(EXIT_SUCCESS);
		}
	}


	cout << "Please enter standard concentration in ppb" << endl;
	while (!(cin >> standardConc)) {
		cerr << "Invalid input. Try again: ";
		cin.clear();
		cin.ignore(numeric_limits<streamsize>::max(), '\n');
	}
	cout << standardConc;

	char selection;
	do {
		cout << endl << "Would you want to calculate recoveries? " << endl
			<< "Y. Yes" << endl
			<< "N. No" << endl;
		cin >> selection;
		if (tolower(selection) == 'y') {
			calculateRecovery = true;
			cout << endl << "Please enter expected sample potency in ppb " << endl;
			while (!(cin >> expectedPotency)) {
				cerr << "Invalid input. Try again: ";
				cin.clear();
				cin.ignore(numeric_limits<streamsize> ::max(), '\n');
			}
		}
	} while (tolower(selection) != 'y' && tolower(selection)!= 'n');

	do {
		cout << endl << "Would you want to calculate average for duplicate injections? " << endl
			<< "Note: Please make sure the duplicate injections have identical names other than injection numbers" <<endl
			<< "For example, <<sampleName INJ1 and sampleName INJ2>> could be associated by the program but" <<endl
			<< " <<sample Name INJ1 and sampleName INJ2>> is not a valid name pair" << endl

			<< "Y. Yes" << endl
			<< "N. No" << endl;
		cin >> selection;
		if (tolower(selection) == 'y') {
			averageMode = true;
		}
		else if (tolower(selection) != 'y' && tolower(selection) != 'n') {
			cout << "Invalid input." <<endl;
		}
	} while (tolower(selection) != 'y' && tolower(selection) != 'n');
	
	calculate_assay(results);

	for (auto result : results) {
		//result.second.peakRatio = result.second.melengestrolArea / result.second.megestrolArea;
		cout << result.first << " " << result.second.sampleType << " " << result.second.megestrolArea << " " << result.second.melengestrolArea << " " <<
			result.second.peakRatio << " "<< result.second.assay<<" " << 100* result.second.recovery<< "%"<< endl;
	}
	
	export_to_xml(results, inputFile);
	return(EXIT_SUCCESS);
}