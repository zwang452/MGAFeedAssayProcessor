#include "utility.h"
#include <iostream>
#include <fstream>
#include <string>
#include <vector>
#include <filesystem>
#include <map>
#include <iterator>
#include <iomanip>
#include <sstream>
#include <math.h>
#include <ctime>
#include "libxl.h"
#pragma warning(disable : 4996)

//Declaring variables; indices of -1 indicates field absent from the input data
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


void help() {
	std::cout << std::endl << "Auto Processor and report generator for MGA Feed Assay" << std::endl << std::endl
		<< "Usage: MGAFeedAssayProcessor [--help|version] infile" << std::endl << std::endl
		<< "Examples:" << std::endl
		<< "  MGAFeedAssayProcessor myfile.txt" << std::endl
		<< "    > reads data from myfile.txt exported from Empower and perfroms calculation " << std::endl << std::endl
		<< "  MGAFeedAssayProcessor --help" << std::endl
		<< "    > displays the help" << std::endl << std::endl
		<< "  rle --version" << std::endl
		<< "    > displays version number in the format d.d.d" << std::endl
		<< std::endl << "Note: this program uses the third-party library libxl to generate output excel files." << std::endl
		<< "Please make sure libxl.dll is placed under the same directory as the program" << std::endl
		<< "Excel file will be generated under the same directory with name format inputFileName + CurrentDate(YYYYMD) + xls." << std::endl
		<< "File with same name under the directory "
		<< "will be overwritten." << std::endl << std::endl
		<< "source code of the program and most updated version can be found under" << std::endl
		<< "https://github.com/zwang452/MGAFeedAssayProcessor" << std::endl;
}

// Function that seperate each value in a line using delimeter and stores them in a std::map
void split(const std::string& s, char delim, std::map<std::string, injection>& results) {
	std::stringstream ss;
	ss.str(s);
	std::string item;
	injection injectionData;
	std::string sampleName;
	std::vector<std::string> elems;
	int indexCount = -1;
	bool isDataRow = false;
	while (getline(ss, item, delim)) {

		//Push each element into a std::vector for access later
		elems.push_back(item);

		//Increase the index by 1 each time an element is loaded
		indexCount++;

		//If a header is matched, set its index
		if (item.find("Sample Type") != std::string::npos) {
			sampleTypeIndex = indexCount;
		}
		else if (item.find("Name") != std::string::npos && item.find("SampleName") == std::string::npos) {
			peakNameIndex = indexCount;
		}
		else if (item.find("RT") != std::string::npos) {
			RTIndex = indexCount;
		}
		else if (item.find("Area") != std::string::npos) {
			areaIndex = indexCount;
		}
		else if (item.find("SampleName") != std::string::npos) {
			sampleNameIndex = indexCount;
		}
		else if (item.find("SampleWeight") != std::string::npos) {
			weightIndex = indexCount;
		}
		else if (item.find("Dilution") != std::string::npos) {
			dilutionIndex = indexCount;
		}
		else {
			isDataRow = true;
		}

		//Sometimes the first row of data file is "Peak results"; Code below ingores that row
		if (item.find("Peak Results") != std::string::npos) {
			isDataRow = false;
		}
	}


	if (isDataRow) {
		sampleName = elems.at(sampleNameIndex);
		auto loc = results.find(sampleName);
		if (loc == results.end()) {
			//If the current injection is not present to the std::map, add it to the std::map as a new key and update its fields
			injectionData.sampleType = elems.at(sampleTypeIndex);
			if (injectionData.sampleType.find("tandard") != std::string::npos) {
				standardInjectionsCount++;
			}
			if (weightIndex != -1) {
				injectionData.weight = stod(elems.at(weightIndex), nullptr);
			}
			if (dilutionIndex != -1) {
				injectionData.dilution = stod(elems.at(dilutionIndex), nullptr);
			}
			if (elems.at(peakNameIndex).find("Melengestrol") != std::string::npos) {
				injectionData.melengestrolArea = stod(elems.at(areaIndex), nullptr);
			}
			else if (elems.at(peakNameIndex).find("Megestrol") != std::string::npos) {
				injectionData.megestrolArea = stod(elems.at(areaIndex), nullptr);
			}
			results.insert(std::pair<std::string, injection>(sampleName, injectionData));
		}
		//If the current injection is present to the std::map, update it 
		else {
			if (elems.at(peakNameIndex).find("Melengestrol") != std::string::npos) {
				(*loc).second.melengestrolArea = stod(elems.at(areaIndex), nullptr);
			}
			else if (elems.at(peakNameIndex).find("Megestrol") != std::string::npos) {
				(*loc).second.megestrolArea = stod(elems.at(areaIndex), nullptr);
			}
		}

	}

}


//Update the sample weights of injections if manual mode is toggled
void mannually_enter_weights(std::map<std::string, injection>& results) {
	for (auto& result : results) {
		if (result.second.sampleType.find("Unknown") != std::string::npos) {
			std::cout << std::endl << "Please enter sample weight in grams:" << std::endl <<
				"sample: " << result.first << std::endl;
			while (!(std::cin >> result.second.weight)) {
				std::cerr << "Invalid input. Try again: ";
				std::cin.clear();
				std::cin.ignore(std::numeric_limits<std::streamsize>::max(), '\n');
			}
		}
	}
}

void manually_enter_dilutions(std::map<std::string, injection>& results) {
	for (auto& result : results) {
		if (result.second.sampleType.find("Unknown") != std::string::npos) {
			std::cout << std::endl << "Please enter dilution factor for sample "
				<< result.first << std::endl;
			while (!(std::cin >> result.second.dilution)) {
				std::cerr << "Invalid input. Try again: ";
				std::cin.clear();
				std::cin.ignore(std::numeric_limits<std::streamsize> ::max(), '\n');
			}

		}
	}
}


//Calculates for calibration with injections marked as standard
void calculate_calibration_curve(std::map<std::string, injection>& results) {
	double standardPeakRatioSum = 0.0;
	for (auto& result : results) {
		if (result.second.sampleType.find("tandard") != std::string::npos) {
			standardPeakRatioSum += result.second.peakRatio;
		}
	}

	standardPeakRatioMean = standardPeakRatioSum / standardInjectionsCount;

	for (auto result : results) {
		if (result.second.sampleType.find("tandard") != std::string::npos) {
			standardStev += pow((result.second.peakRatio - standardPeakRatioMean), 2);
		}
	}
	standardStev /= (standardInjectionsCount - 1);
	standardStev = sqrt(standardStev);
	standardRSD = standardStev / standardPeakRatioMean;
}

//Calculate assay results in unit of ppb using the standard calibration curve
void calculate_assay(std::map<std::string, injection>& results) {
	for (auto& result : results) {
		if (result.second.sampleType.find("Unknown") != std::string::npos) {
			result.second.assay = (result.second.peakRatio / standardPeakRatioMean) * standardConc * (result.second.dilution / result.second.weight);
			if (calculateRecovery == true) {
				result.second.recovery = result.second.assay / expectedPotency;
			}
		}
	}
}

//Export data to excel using libxl
void export_to_xml(std::map<std::string, injection>& results, std::string& inputFile) {

	time_t now = time(0);
	char buffer[26];
	char* stringNow = ctime(&now);
	wchar_t wcharNow[100];
	mbstowcs(wcharNow, stringNow, 100);


	//Format date
	tm* ltm = localtime(&now);
	int month = 1 + ltm->tm_mon;
	int day = ltm->tm_mday;
	int year = 1900 + ltm->tm_year;
	std::string outputNamestring = inputFile.substr(0, inputFile.length() - 3) + std::to_string(year) + std::to_string(month) + std::to_string(day) + ".xls  ";
	std::wstring outputNameWstring = std::wstring(outputNamestring.begin(), outputNamestring.end());
	const wchar_t* outputWchar = outputNameWstring.c_str();
	//Create new instance of xml file
	libxl::Book* book = xlCreateBook();

	libxl::Format* percentage = book->addFormat();
	percentage->setNumFormat(libxl::NUMFORMAT_PERCENT_D2);

	libxl::Font* redFont = book->addFont();
	redFont->setColor(libxl::COLOR_RED);
	libxl::Format* redFontFormat = book->addFormat();
	redFontFormat->setFont(redFont);
	redFontFormat->setBorderColor(libxl::COLOR_RED);

	libxl::Font* greenFont = book->addFont();
	greenFont->setColor(libxl::COLOR_GREEN);
	libxl::Format* greenFontFormat = book->addFormat();
	greenFontFormat->setFont(greenFont);
	greenFontFormat->setBorderColor(libxl::COLOR_GREEN);


	if (book) {
		libxl::Sheet* sheet = book->addSheet(L"Sheet1");
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
				if (result.second.sampleType.find("tandard") != std::string::npos) {
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
			if ((int)(standardRSD * 100) <= 10) {
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
				if (result.second.sampleType.find("Unknown") != std::string::npos) {

					sheet->writeStr(row, col++, std::wstring(result.first.begin(), result.first.end()).c_str());
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
						sheet->writeNum(row, col++, (firstAssay + result.second.assay) / 2.0);
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
				std::cout << "Data output to excel file named " + outputNamestring + " successfully!" << std::endl;
			}
			else
			{
				std::cout << book->errorMessage() << std::endl
					<< "Please check there is no file named " + outputNamestring + "open under the same directory" << std::endl
					<< "No excel file generated" << std::endl;
			}
			book->release();
		}
	}
}