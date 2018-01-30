import json
import requests

#import Image
#from io import BytesIO

data =[]

##  Reads in json
with open ('showcase_backup/projects/all_projects.json') as p:
	for line in p:
		data.append(json.loads(line))
p.close()

##  Loops through image refrences
for i in data:
	#print(i['image_ref'])
	digest = i['image_ref']
	r = requests.get('http://localhost:4200/_blobs/project_images/%s' % digest)
	
	##  Creates and saves image using PIL(pillow) library that currently is not installed.   
	#backup_image=Image.open(BytesIO(r.content))
	#backup_image.save("%s.jpg" % digest)

	##  Creates and saves image.	
	image_backup = open("showcase_backup/images/%s.jpg" % digest, "wb")
	image_backup.write(r.content)
	image_backup.close()



