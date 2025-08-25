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
import pdfplumber


st.set_page_config(layout="wide")

st.markdown(f"<h1 style='text-align: center;'>WELCOME TO OCR PAJAK<br>LAUTAN LUAS</h1>", unsafe_allow_html=True)

# if "page" not in st.session_state:
#     st.session_state["page"] = 0

# if st.session_state["page"] == 0 :
#     tombol_mulai = st.button(type="primary", label='Mulai')
#     if tombol_mulai:
#         st.session_state.page = 1
#         st.rerun()


# if st.session_state["page"] == 1:
    
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
        # st.write(path_to_pdf)

        file_path_pdf = os.listdir(path_to_pdf)
        # st.write(file_path_pdf)

        file_count = len(file_path_pdf)



        saved_directory = os.path.join(os.getcwd(), 'saved_image' + ' ' + os.path.splitext(user_input_folder.name)[0])

        if not os.path.exists(saved_directory):
            os.makedirs(saved_directory)



        with st.spinner("Wait for it..."):
            with st.empty():
                for i in range(len(file_path_pdf)):
                
                    st.write("Converting "+str(i+1)+"/"+str(file_count))
                    images = convert_from_path(os.path.join(path_to_pdf,file_path_pdf[i]), 500)
                    for j, image in enumerate(images):
                        fname = os.path.join(saved_directory, str(file_path_pdf[i])[:-4]+'.jpg')
                        image.save(fname, "JPEG")
            
            time.sleep(0.5)
        st.success("File converted successfully!")


        nama_kolom = {
            "NOMOR": [],
            "MASA PAJAK": [],
            "SIFAT PEMOTONGAN DAN/ATAU PEMUNGUTAN PPh": [],
            "STATUS BUKTI PEMOTONGAN / PEMUNGUTAN": [],
            "B.2 Jenis PPh": [],
            "KODE OBJEK PAJAK": [],
            # "OBJEK PAJAK": [],
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
            # "DPP converted": [],
            # "PAJAK PENGHASILAN converted": [],
            # "TARIF converted": []
            }
        df_all_data_extracted_combined = pd.DataFrame(nama_kolom)

        # st.write(os.listdir(os.getcwd()))
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

                    # ONLY FOR NOMOR
                    
                    # Open the PDF file

                    current_filename=image_path_in_colab[(len(saved_directory)+1):][:-4] + str(".pdf")
                    
                    # st.write(current_filename)
                    # st.write(os.path.join(path_to_pdf, current_filename))
                    reader = PdfReader(os.path.join(path_to_pdf, current_filename))
                    a=[]
                    # Iterate through pages and extract text
                    extracted_text = ""
                    for page in reader.pages:
                        extracted_text += page.extract_text()
                    a = [extracted_text]
                    
                    text_for_nomor = ""
                    text_for_b2 = ""
                    text_for_b8_jenisdokumen = ""
                    text_for_b8_tanggal = ""
                    text_for_b9 = ""
                    text_for_c1 = ""
                    text_for_c2 = ""
                    text_for_c3 = ""
                    text_for_c4 = ""
                    
                    text_for_nomor = re.findall('(?<=PEMUNGUTAN PPh PEMUNGUTAN\n)[^ ]+', a[0])[0]
                    text_for_b2 = re.findall('(?<=B.2 Jenis PPh : )[^ ].*', a[0])[0]
                    text_for_b8_jenisdokumen = re.findall('(?<=B.8 Dokumen Dasar Bukti\nPemotongan dan/atau\nPemungutan PPh Unifikasi\natau Dasar Pemberian\nFasilitas\nJenis Dokumen : )[^ ].*(?= Tanggal)', a[0])[0]
                    text_for_b8_tanggal = re.findall('(?<=Tanggal : )[^ ].*', re.findall('(?<=B.8 Dokumen Dasar Bukti\nPemotongan dan/atau\nPemungutan PPh Unifikasi\natau Dasar Pemberian\nFasilitas\nJenis Dokumen : )[^ ].*', a[0])[0])[0]
                    text_for_b9 = re.findall('(?<=B.9  Nomor Dokumen : )[^ ].*', a[0])[0]
                    text_for_c1 = re.findall('(?<=C.1 NPWP / NIK : )[^ ].*', a[0])[0]
                    text_for_c2 = re.findall('(?<=SUBUNIT ORGANISASI\n: )[^ ].*', a[0])[0]
                    text_for_c3 = re.findall('(?<=C.3 NAMA PEMOTONG DAN/ATAU PEMUNGUT\nPPh\n: )[^ ].*', a[0])[0]
                    text_for_c4 = re.findall('(?<=C.4 TANGGAL : )[^ ].*', a[0])[0]


                    extracted_text_b567 = ""

                    with pdfplumber.open(os.path.join(path_to_pdf, current_filename)) as pdf:
                        for page in pdf.pages:
                            # Use a corrected bounding box (x0, y0, x1, y1)
                            # (left, bottom, right, top)
                            cropped = page.within_bbox((int(11/20.35*595), int((20.35-10)/20.35*595), int(19.45/20.35*595), int((20.35-9.35)/20.35*595)))
                            text = cropped.extract_text()
                            if text:
                                extracted_text_b567 += text.strip()
                    b567 = re.split(r'\s+', extracted_text_b567)
                    text_for_b5 = b567[0]
                    text_for_b6 = b567[1]
                    text_for_b7 = b567[2]
                    # st.write(b567)
                    # st.write(text_for_b5)
                    # st.write(text_for_b6)
                    # st.write(text_for_b7)

        
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
                        # [1.5, 5.3, 3.8, 4.2],
                        [6.1, 9.9, 3.8, 4.2],
                        [11, 14, 3.8, 4.2],
                        [15.5, 19, 3.8, 4.2],
                        # [3.35, 5, 8.2, 8.8],
                        [2, 5.2, 10.35, 10.8],
                        # [5.4, 10.2, 10.35, 10.8],
                        # [11, 13.2, 10.35, 10.8],
                        # [14, 15, 10.35, 10.8],
                        # [15.7, 19.4, 10.35, 10.8],
                        # [9, 12, 11.2, 12],
                        # [14.2, 17, 11.2, 12],
                        # [9, 13, 13.2, 13.7],
                        # [7.5, 20, 15.3, 15.8],
                        # [7.5, 20, 15.8, 16.7],
                        # [7.5, 20, 16.7, 17.7],
                        # [7.5, 20, 17.7, 18.2]
                    ]
        
        
        
                    nama_kolom = {
                        "NOMOR": [],
                        "MASA PAJAK": [],
                        "SIFAT PEMOTONGAN DAN/ATAU PEMUNGUTAN PPh": [],
                        "STATUS BUKTI PEMOTONGAN / PEMUNGUTAN": [],
                        "B.2 Jenis PPh": [],
                        "KODE OBJEK PAJAK": [],
                        # "OBJEK PAJAK": [],
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
                        # "DPP converted": [],
                        # "PAJAK PENGHASILAN converted": [],
                        # "TARIF converted": []
                        }
                    df_all_data = pd.DataFrame(nama_kolom)
        
                    

                    def extract_text(image=img, coordinates=coordinates, all_data=df_all_data, text_for_nomor=text_for_nomor):
                        extracted=[]
                        for i in range(len(coordinates)):
        
                            x_start, x_end, y_start, y_end = region_of_interest(coordinates[i])
        
                            cropped_img = img[y_start:y_end, x_start:x_end]
                            cropped_img_bigger = cv2.copyMakeBorder(cropped_img, 200, 200, 200, 200, cv2.BORDER_CONSTANT, value=(255, 255, 255))
        
                            extractedInformation = pytesseract.image_to_string(cropped_img_bigger).strip()
                            # extractedInformation = pytesseract.image_to_string(cropped_img).strip()
                            # extractedInformation = reader.readtext(cropped_img_bigger, detail=0)
        
                            extracted.append(extractedInformation)

                        
                            
                        
                        new_row = pd.DataFrame({
                                            "NOMOR": [text_for_nomor],
                                            "MASA PAJAK": [extracted[0]],
                                            "SIFAT PEMOTONGAN DAN/ATAU PEMUNGUTAN PPh": [extracted[1]],
                                            "STATUS BUKTI PEMOTONGAN / PEMUNGUTAN": [extracted[2]],
                                            "B.2 Jenis PPh": [text_for_b2],
                                            "KODE OBJEK PAJAK": [extracted[3]],
                                            # "OBJEK PAJAK": [],
                                            "DPP": [text_for_b5],
                                            "TARIF": [text_for_b6],
                                            "PAJAK PENGHASILAN": [text_for_b7],
                                            "B.8 Jenis Dokumen": [text_for_b8_jenisdokumen],
                                            "B.8 Tanggal": [text_for_b8_tanggal],
                                            "B.9 Nomor Dokumen": [text_for_b9],
                                            "C.1 NPWP / NIK": [text_for_c1],
                                            "C.2 NOMOR IDENTITAS TEMPAT KEGIATAN USAHA (NITKU) / SUBUNIT ORGANISASI": [text_for_c2],
                                            "C.3 NAMA PEMOTONG DAN/ATAU PEMUNGUT": [text_for_c3],
                                            "C.4 TANGGAL": [text_for_c4],
                                            "Nama File": [image_path_in_colab[(len(saved_directory)+1):][:-4]],
                                            # "DPP converted": [float(re.sub(r"[\s.]" , "", extracted[4]))],
                                            # "PAJAK PENGHASILAN converted": [float(re.sub(r"[\s.]" , "", extracted[5]))],
                                            # "TARIF converted": [round(float(re.sub(r"[\s.]" , "", extracted[5]))/float(re.sub(r"[\s.]" , "", extracted[4]))*100, 2)]
                                        })
                        df_all_data_extracted = pd.concat([df_all_data, new_row]).reset_index(drop=True)
                        return(df_all_data_extracted)
        
                    df_all_data_extracted = extract_text(image=img, coordinates=coordinates, all_data=df_all_data, text_for_nomor=text_for_nomor)
        
                    df_all_data_extracted_combined = pd.concat([df_all_data_extracted_combined, df_all_data_extracted]).reset_index(drop=True)

                    j+=1
            

            
                # time.sleep(0.5)

            with st.spinner("Preparing to show some samples of data ..."):
                st.dataframe(df_all_data_extracted_combined.head(5))

            with st.spinner("Preparing for data to be downloaded ..."):
                output = BytesIO()
    
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer: 
                    df_download = df_all_data_extracted_combined.to_excel(writer)

                output_excel_file_name = 'result' + ' ' + str(os.path.splitext(user_input_folder.name)[0]) + '.xlsx'
        
                button_clicked = st.download_button(label=':cloud: Download result', type="secondary", data=output.getvalue(),file_name=output_excel_file_name)

        end_datetime = datetime.datetime.now()
        time_difference = end_datetime - current_datetime
    
        st.write(f"Running Time: {time_difference}")

            
        # tombol_ulangi = st.button(type="primary", label='Ulangi dari awal')
        # if tombol_ulangi:
        #     st.session_state.page = 0
        #     st.rerun()
    

        
    else:
        st.warning('You need to upload zip type file')
    

else :
    st.error("You have to upload pdf folder in the sidebar")










































































