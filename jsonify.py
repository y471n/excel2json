import numpy as np
from xlrd import open_workbook
import re
import fileinput

START_ROW = 1
START_COLUMN = 0
END_COLUMN = 3
INPUT_EXCEL = 'album_info.xls'

wb = open_workbook(INPUT_EXCEL,'UTF-8')

OUTPUT_FILE_NAME = 'album_info.json'
json_file = open(OUTPUT_FILE_NAME,'w')

def main():

	i = 1
	
	for s in wb.sheets():
		print 'In Sheet: ', s.name.capitalize()
		END_COLUMN = s.ncols																				 # End Column changed here
		
		for row in range(START_ROW,s.nrows):

			json_file_contents = "\n{ \n\t \"id\" : \"" + str(i) + "\", "                                    #  Change here for index 

			for col in range(START_COLUMN,END_COLUMN):

				cell_value = s.cell(row,col).value				
				title = ''
				album = ''
				url = ''
				if (s.cell(0, col).value.encode('UTF-8') == 'title'):										 # Title value got here
					title = s.cell(row, col).value.encode('UTF-8')
					print title

				if (s.cell(0, col).value.encode('UTF-8') == 'album'):										 # Album value got here
					album = s.cell(row, col).value.encode('UTF-8')
					print album

				if (s.cell(0,col).value.encode('UTF-8') == 'mobile_url'):									 # Mobile Url got here
					url = s.cell(row, col).value.encode('UTF-8')

				if len(title)!=0:																			 # Last item is last column to enter json
					json_file_contents = json_file_contents + "\n\t \"title\" : \"" + title + "\","          # Title put here
				if len(album)!=0:
				    json_file_contents = json_file_contents + "\n\t \"album\" : \"" + album + "\"\n}"		 # Album put here - < Last item >
				    if row != s.nrows-1:
				    	json_file_contents = json_file_contents + ','
				if len(url)!=0:
					json_file_contents = json_file_contents + "\n\t \"url\" : \"" + url + "\","              # Mobile url put here


			i = i + 1

			
			json_file.write(json_file_contents)
		
		
if __name__ == '__main__':
	main()
