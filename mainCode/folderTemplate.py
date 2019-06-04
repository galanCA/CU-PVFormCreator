import os


def main():
	central_path = input("Folder Path: ")
	date_use = input("Date to be in the title (yyyy_mm_dd): ")
	#central_path = "C:\Users\Cesar Workdesk\Documents\IRL\PV Forms\Heckman"

	dirName = "/PV Form " + date_use.split(" ")[0]
	PV = "/PV"
	receipts = "/Receipts"

	# Create the main folder
	os.mkdir(central_path+dirName)

	# Create subfolders
	os.mkdir(central_path+dirName + PV)
	os.mkdir(central_path+dirName + receipts)

if __name__ == '__main__':
	main()