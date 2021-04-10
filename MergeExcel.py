#/bin/usr/python

import pandas as pd;
import sys
import os
import glob

def mergeExcel(files,outfile):
	allFrames=[];
	for f in files:
		df = pd.read_excel(f,index_col=0);
		allFrames.append(df);
	merged_df = pd.concat(allFrames);
	print(merged_df);
	merged_df.to_excel(outfile,sheet_name="AllRecords");
	print("Successfully written merge result to "+outfile);

if __name__=="__main__":
	if len(sys.argv) == 1:
		print("Usage: python MergeExcel.py [Source Directory] [OuputFile]");
	else:
		filesList=[];
		os.chdir(sys.argv[1]);
		for f in glob.glob("*.xls"):
			filesList.append(f);
		for f in glob.glob("*.xlsx"):
			filesList.append(f);
		print("Merging total files :"+str(len(filesList)));
		mergeExcel(filesList,sys.argv[2]);
