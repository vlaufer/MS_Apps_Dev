from PIL import Image

def standardize_picture_files(path_2_pictures, picture_width_inches, picture_file_type):

	picture_test_fname=picture_dir + "/bluegene-hi-res" + ".jpg"
	output_fname=output_dir + "/formatted_" + "bluegene-hi-res" + "_formatted" + ".jpg"

#	im1 = Image.open(r'path where the JPG is stored\file name.jpg')
#	im1.save(r'path where the PNG will be stored\new file name.

	im1 = Image.open(picture_test_fname)
	im1.save(output_fname)



base_dir="/Users/vincentlaufer/Desktop/Consultancies/Laufer_D/Automated_Excel_Functions"
script_dir=base_dir + "/" + "Code"
data_dir=base_dir + "/" +"Data"
picture_dir=data_dir + "/" +"Source_Pictures"
output_dir=base_dir + "/" +"Results"

# Picture file parameters
picture_width_inches=4
picture_file_type="jpg"
compression_format="lzw"

standardize_picture_files(picture_dir, picture_width_inches, picture_file_type)


for root, dirs, files in os.walk(folder_path):
	for file in files:
		upload_file(client, folder.id, os.path.join(folder_path, file))
		for d in dirs:
			upload_to_box(client, folder.id, os.path.join(folder_path, d), d)
		break;