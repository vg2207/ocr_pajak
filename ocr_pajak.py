import streamlit as st
import os
from pathlib import Path
import pandas as pd
import xlsxwriter
import glob
import shutil
import zipfile
import pytesseract
import cv2
from pdf2image import convert_from_path
import time
from io import BytesIO
import datetime
# import easyocr
import re
from pypdf import PdfReader


st.set_page_config(layout="wide")





user_input_folder = st.file_uploader("Upload pdf folder", type=['zip'], accept_multiple_files=False, key='file_uploader')




if user_input_folder is not None:
    if user_input_folder.name.endswith('.zip'):
        
        current_datetime = datetime.datetime.now()
        
        target_path = os.path.join(os.getcwd(), os.path.splitext(user_input_folder.name)[0])
        if os.path.exists(target_path) == False:
            os.mkdir(target_path)
        with zipfile.ZipFile(user_input_folder, 'r') as z:
            z.extractall(target_path)
        st.success('Folder Uploaded Successfully!')

        path_to_pdf = os.path.join(target_path, str(os.listdir(target_path)[0]))

        file_path_pdf = os.listdir(path_to_pdf)

        file_count = len(file_path_pdf)


        saved_directory = 'saved_image'
        if not os.path.exists(saved_directory):
            os.makedirs(saved_directory)


        with st.spinner("Wait for it..."):
            with st.empty():
                for i in range(len(file_path_pdf)):
                
                    st.write("Converting "+str(i+1)+"/"+str(file_count))
                    images = convert_from_path(os.path.join(path_to_pdf,file_path_pdf[i]), 500)
                    for j, image in enumerate(images):
                        fname = os.path.join(saved_directory, str(file_path_pdf[i])[:-4]+'.jpg')
            
            time.sleep(0.5)
        st.success("File converted successfully!")


        nama_kolom = {
            "NOMOR": [],
            "MASA PAJAK": [],
            "SIFAT PEMOTONGAN DAN/ATAU PEMUNGUTAN PPh": [],
            "STATUS BUKTI PEMOTONGAN / PEMUNGUTAN": [],
            "B.2 Jenis PPh": [],
            "KODE OBJEK PAJAK": [],
            "OBJEK PAJAK": [],
            "DPP": [],
            "TARIF": [],
            "PAJAK PENGHASILAN": [],
            "B.8 Jenis Dokumen": [],
            "B.8 Tanggal": [],
            "B.9 Nomor Dokumen": [],
            "C.1 NPWP / NIK": [],
            "C.2 NOMOR IDENTITAS TEMPAT KEGIATAN USAHA (NITKU) / SUBUNIT ORGANISASI": [],
            "C.3 NAMA PEMOTONG DAN/ATAU PEMUNGUT": [],
            "C.4 TANGGAL": [],
            "DPP converted": [],
            "PAJAK PENGHASILAN converted": [],
            "TARIF converted": []
            }
        df_all_data_extracted_combined = pd.DataFrame(nama_kolom)

        # st.write(os.listdir(saved_directory))
        # reader = easyocr.Reader(['id','en'], gpu=False) # this needs to run only once to load the model into memory
        
        with st.spinner("Wait for it..."):
            with st.empty():
                j=0
                for image_path_in_colab in glob.glob(str(os.path.join(saved_directory+"/*.jpg"))):
                    
                    st.write("Processing "+str(j+1)+"/"+str(file_count))
                    img = cv2.imread(image_path_in_colab, cv2.IMREAD_GRAYSCALE)
        
                    # print(img.shape[1])
                    # Define the region of interest (ROI) - arbitrary coordinates
        
                    def region_of_interest(coordinate):
                        x1 = coordinate[0]
                        x2 = coordinate[1]
                        y1 = coordinate[2]
                        y2 = coordinate[3]
                        x_start = int(x1/20.35*4134)
                        x_end = int(x2/20.35*4134)
                        y_start = int(y1/20.35*4134)
                        y_end = int(y2/20.35*4134)
                        
                        return(x_start, x_end, y_start, y_end)
        
        
                    coordinates = [
                        [1.5, 5.3, 3.8, 4.2],
                        [6.1, 9.9, 3.8, 4.2],
                        [11, 14, 3.8, 4.2],
                        [15.5, 19, 3.8, 4.2],
                        [3.35, 5, 8.2, 8.8],
                        [2, 5.2, 10.35, 10.8],
                        [5.4, 10.2, 10.35, 10.8],
                        [11, 13.2, 10.35, 10.8],
                        [14, 15, 10.35, 10.8],
                        [15.7, 19.4, 10.35, 10.8],
                        [9, 12, 11.2, 12],
                        [14.2, 17, 11.2, 12],
                        [9, 13, 13.2, 13.7],
                        [7.5, 20, 15.3, 15.8],
                        [7.5, 20, 15.8, 16.7],
                        [7.5, 20, 16.7, 17.7],
                        [7.5, 20, 17.7, 18.2]]
        
        
        
                    nama_kolom = {
                        "NOMOR": [],
                        "MASA PAJAK": [],
                        "SIFAT PEMOTONGAN DAN/ATAU PEMUNGUTAN PPh": [],
                        "STATUS BUKTI PEMOTONGAN / PEMUNGUTAN": [],
                        "B.2 Jenis PPh": [],
                        "KODE OBJEK PAJAK": [],
                        "OBJEK PAJAK": [],
                        "DPP": [],
                        "TARIF": [],
                        "PAJAK PENGHASILAN": [],
                        "B.8 Jenis Dokumen": [],
                        "B.8 Tanggal": [],
                        "B.9 Nomor Dokumen": [],
                        "C.1 NPWP / NIK": [],
                        "C.2 NOMOR IDENTITAS TEMPAT KEGIATAN USAHA (NITKU) / SUBUNIT ORGANISASI": [],
                        "C.3 NAMA PEMOTONG DAN/ATAU PEMUNGUT": [],
                        "C.4 TANGGAL": [],
                        "Nama File": [],
                        "DPP converted": [],
                        "PAJAK PENGHASILAN converted": [],
                        "TARIF converted": []
                        }
                    df_all_data = pd.DataFrame(nama_kolom)
        
                    

                    def extract_text(image=img, coordinates=coordinates, all_data=df_all_data):
                        extracted=[]
                        for i in range(len(coordinates)):
        
                            x_start, x_end, y_start, y_end = region_of_interest(coordinates[i])
        
                            cropped_img = img[y_start:y_end, x_start:x_end]
                            # cropped_img_bigger = cv2.copyMakeBorder(cropped_img, 200, 200, 200, 200, cv2.BORDER_CONSTANT, value=(255, 255, 255))
        
                            # extractedInformation = pytesseract.image_to_string(cropped_img_bigger).strip()
                            extractedInformation = pytesseract.image_to_string(cropped_img).strip()
                            # extractedInformation = reader.readtext(cropped_img_bigger, detail=0)
        
                            extracted.append(extractedInformation)

                        # ONLY FOR NOMOR
                        for j in range(len(file_path_pdf)):
                            # Open the PDF file
                            reader = PdfReader(os.path.join(path_to_pdf,file_path_pdf[j])
                            
                            # Iterate through pages and extract text
                            extracted_text = ""
                            for page in reader.pages:
                                extracted_text += page.extract_text()
                            a = [extracted_text]
                            text_for_nomor = re.findall('(?<=PEMUNGUTAN PPh PEMUNGUTAN\n)[^ ]+', a[0])
                            
                        
                        new_row = pd.DataFrame({
                                            # "NOMOR": [extracted[0]],
                                            "NOMOR": [text_for_nomor],
                                            "MASA PAJAK": [extracted[1]],
                                            "SIFAT PEMOTONGAN DAN/ATAU PEMUNGUTAN PPh": [extracted[2]],
                                            "STATUS BUKTI PEMOTONGAN / PEMUNGUTAN": [extracted[3]],
                                            "B.2 Jenis PPh": [extracted[4]],
                                            "KODE OBJEK PAJAK": [extracted[5]],
                                            "OBJEK PAJAK": [extracted[6]],
                                            "DPP": [extracted[7]],
                                            "TARIF": [extracted[8]],
                                            "PAJAK PENGHASILAN": [extracted[9]],
                                            "B.8 Jenis Dokumen": [extracted[10]],
                                            "B.8 Tanggal": [extracted[11]],
                                            "B.9 Nomor Dokumen": [extracted[12]],
                                            "C.1 NPWP / NIK": [extracted[13]],
                                            "C.2 NOMOR IDENTITAS TEMPAT KEGIATAN USAHA (NITKU) / SUBUNIT ORGANISASI": [extracted[14]],
                                            "C.3 NAMA PEMOTONG DAN/ATAU PEMUNGUT": [extracted[15]],
                                            "C.4 TANGGAL": [extracted[16]],
                                            "Nama File": [image_path_in_colab[12:][:-4]],
                                            "DPP converted": [float(extracted[7].replace(".",""))],
                                            "PAJAK PENGHASILAN converted": [float(extracted[9].replace(".",""))],
                                            "TARIF converted": [round(float(extracted[9].replace(".",""))/float(extracted[7].replace(".",""))*100, 2)]
                                        })
                        df_all_data_extracted = pd.concat([df_all_data, new_row]).reset_index(drop=True)
                        return(df_all_data_extracted)
        
                    df_all_data_extracted = extract_text(image=img, coordinates=coordinates, all_data=df_all_data)
        
                    df_all_data_extracted_combined = pd.concat([df_all_data_extracted_combined, df_all_data_extracted]).reset_index(drop=True)

                    j+=1
            

            
                # time.sleep(0.5)

            with st.spinner("Preparing to show some samples of data ..."):
                st.dataframe(df_all_data_extracted_combined.head(5))

            with st.spinner("Preparing for data to be downloaded ..."):
                output = BytesIO()
    
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer: 
                    df_download = df_all_data_extracted_combined.to_excel(writer)
        
                button_clicked = st.download_button(label=':cloud: Download result', type="secondary", data=output.getvalue(),file_name='result.xlsx')

        end_datetime = datetime.datetime.now()
        time_difference = end_datetime - current_datetime
    
        st.write(f"Running Time: {time_difference}")


        
    else:
        st.warning('You need to upload zip type file')
    

else :
    st.error("You have to upload pdf folder in the sidebar")



















