#pragma once
#include <string>
#include <map>



extern int sampleTypeIndex;
extern int sampleNameIndex;
extern int peakNameIndex;
extern int RTIndex;
extern int areaIndex;
extern int weightIndex;
extern int dilutionIndex;
extern double expectedPotency;
extern size_t standardInjectionsCount;
extern double standardPeakRatioMean;
extern double standardStev;
extern double standardRSD;
extern bool calculateRecovery;
extern double standardConc;
extern bool averageMode;


struct injection {
	std::string sampleType;
	double melengestrolArea = 0;
	double megestrolArea = 0;
	double peakRatio = 0.0;
	double dilution = 1.0;
	double weight = 1.0;
	double assay = 0.0;
	double recovery = 0.0;
};

void help();
void split(const std::string& s, char delim, std::map<std::string, injection>& results);
void mannually_enter_weights(std::map<std::string, injection>& results);
void manually_enter_dilutions(std::map<std::string, injection>& results);
void calculate_calibration_curve(std::map<std::string, injection>& results);
void calculate_assay(std::map<std::string, injection>& results);
void export_to_xml(std::map<std::string, injection>& results, std::string& inputFile);