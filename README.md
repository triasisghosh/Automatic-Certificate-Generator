# Automatic-Certificate-Generator
This simple program is a tkinter GUI directed solution for making certificates for mass participation in events, webinar, exams or any oher relevent fields.

# How to use:
1. A CSV file containing records of recipients, a empty template in JPG/JPEG/PNG format, a destination folder whwere the certificates will be saved after creation is all what you need to upload in specifie fields.
2. Select the row number from where the list of recipents starts and the key column (the column having unique and identifiable value for each records) in the CSV file. The value of key-column will be used name the certificate file.
3. By default there are some fields which i to be printed in certificate. Any extra field can be added by selecting the source for the field. Also any field can be disabled or deleted if not required.
4. There are mainly three types of sources for the fields, which are <br/>
  *a. a particular column for every record(Eg. phone number, name etc)<br/>
  b. a constant text(Eg. instituition name, appreciation messages etc.)<br/>
  c. date<br/>
  d. constant image(Eg. Logo, signature etc. To be uploaded in specified field).<br/>*
5. Style, Colour, Size can be adjusted for the text type things(i.e a,b and c of last point) to be printed.
6. Date can be chosen from the calender wiget which can be accessesed by specified button. also the date formats are fully customisable.
7. Position of every fields is to be selected by clicking desired position on the template through the given button.
8. A preview of the sample certificate can be obtained by clicking the 'PREVIEW' button. After being satisfied by the settings, 'GENERATE' button will generate all the certificates in PNG format to the destination folder.

# Required libraries and files:
- Python 3
- Tkinter
- OpenCV
- Pillow or PIL(Python Image Processing)
- A white image of resolution 100x30 named as 'preview.png' should be present in the root of program file for preview during font selection.

/***A detailed information of the program is enclosed in PDF format in this repository**/*

# Reach me
* instagram @triasisghosh
* linkedin at https://www.linkedin.com/in/triasis-ghosh-322b27201
