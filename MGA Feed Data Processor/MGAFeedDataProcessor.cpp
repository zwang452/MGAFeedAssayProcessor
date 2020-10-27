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
#include <filesystem>
#include <map>
#include "Utility.h"


using namespace std;
using namespace std::filesystem;



const string VERSION = "1.0.1";

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

	//Data structure that holds injection data from csv
	map<string, injection> results;
	
	//Read injection data line by line
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

	//Previewing peak ratios
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


	//Calculate for calibration is standard injections are present
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


	//Ask for sample weights if they are missing from input data file
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

	//Ask for sample dilutions if they are missing from input data file
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

	//Asks for pre-determined standard concentration 
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
	
	//Calculate from calibration curve
	calculate_assay(results);

	for (auto result : results) {
		cout << result.first << " " << result.second.sampleType << " " << result.second.megestrolArea << " " << result.second.melengestrolArea << " " <<
			result.second.peakRatio << " "<< result.second.assay<<" " << 100* result.second.recovery<< "%"<< endl;
	}
	
	//Export to excel using libxl library
	export_to_xml(results, inputFile);
	return(EXIT_SUCCESS);
}