
################
# Removes excess information from the Skinect .ply file 
# to make it readable by the Lidar Viewer
################

import sys

def main():
	try:
		Cat(sys.argv[1])
	except IndexError:
		print('Please select a file to convert')

def Cat(filename):
	#filename ='C:/Users/pdbeard/Desktop/Dave.ply'
	plyfile = open(filename, "r")
	lidar = open(filename+'.lidar.txt','w')
	for line in plyfile:
		a = line.split()        # Splits each line into its own list
		try:
			float(a[0])         # Checks to see if first item in list is a number
			a[3:6]=[]           # Removes middle 3 numbers.
			
			spaceFormat = ' '.join(a)         # Joins the list into one string  
			lidar.write(spaceFormat + '\n')
			c = True
		except ValueError:
			c = False
	lidar.close
	plyfile.close
if __name__ == '__main__':
	main()

		
				
		
