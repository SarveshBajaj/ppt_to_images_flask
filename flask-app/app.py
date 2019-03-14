import os
import pdf2image
from flask import Flask, Response, request, send_file
from requests_toolbelt import MultipartEncoder
from PIL import Image
from os import listdir
from os.path import isfile, join


ALLOWED_EXTENSIONS = {'ppt', 'pptx', 'odt'}

app = Flask(__name__)

def allowed_extensions(filename):
    return '.' in filename and \
        filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def hello():
    return('hello')

@app.route('/convert',methods=['POST'])
# accepts a ppt file from request, checks if its ppt using allowed_extensions
# then convert ppt to images, returns a list of urls to access the iamges

#install libreoffice, unoconv, pillow, pdf2image
def convertppt_to_images():
    if request.files:
        file = request.files.get('pptfile')
        if file and allowed_extensions(file.filename):
            file.save('ppts/'+file.filename)
            command = "unoconv -f pdf ppts/"+ file.filename
            os.system(command)
            file_pdf = file.filename.split('.')[0]
            if os.path.exists("ppts/"+file_pdf):
                os.system('rm -rf ppts/'+file_pdf)
            os.system('mkdir ppts/'+file_pdf)
            images = pdf2image.convert_from_path('ppts/'+file_pdf+'.pdf', output_folder = 'ppts/'+file_pdf)
            count = 1
            for image in images:
                image.save('ppts/'+file_pdf+'/'+file_pdf+str(count)+'.jpg', 'JPEG')
                count+=1
            # import pdb;pdb.set_trace()
            for file in os.listdir('ppts/'+file_pdf):
                if file.endswith(".ppm"):
                    os.remove('ppts/'+file_pdf+'/'+file)
            onlyfiles = [f for f in listdir('ppts/'+file_pdf) if isfile(join('ppts/'+file_pdf, f))]
            return 'file uploaded'

@app.route('/getimage',methods=['GET'])
# get filename from request params and then send that file using send_file
# filename format : pptname(folder)/imagename(slide number)
def sendimage():
    path = request.args.get('path')
    name = path.split('/')[1]
    if os.path.exists("ppts/"+path):
        return send_file('ppts/'+path, attachment_filename=name)
    return Response('NOT FOUND',status = 404)


#404 error something

if __name__ == '__main__':
    app.run(debug=True);
