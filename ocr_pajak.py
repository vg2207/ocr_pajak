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



st.set_page_config(layout="wide")





user_input_folder = st.file_uploader("Upload pdf folder", type=['zip'], accept_multiple_files=False, key='file_uploader')




if user_input_folder is not None:
    if user_input_folder.name.endswith('.zip'):
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
        # st.write(file_count)

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
                        image.save(fname, "JPEG")
            time.sleep(0.5)
        st.success("File converted!")


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
            "C.4 TANGGAL": []
            }
        df_all_data_extracted_combined = pd.DataFrame(nama_kolom)
        st.write(os.path.join(saved_directory+"/*.jpg"))

        for image_path_in_colab in glob.glob(str(os.path.join(saved_directory+"/*.jpg"))):
            st.write(image_path_in_colab)
            img = cv2.imread(image_path_in_colab, cv2.IMREAD_GRAYSCALE)

            print(img.shape[1])
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
                "C.4 TANGGAL": []
                }
            df_all_data = pd.DataFrame(nama_kolom)


            def extract_text(image=img, coordinates=coordinates, all_data=df_all_data):
                extracted=[]
                for i in range(len(coordinates)):

                    x_start, x_end, y_start, y_end = region_of_interest(coordinates[i])

                    cropped_img = img[y_start:y_end, x_start:x_end]
                    cropped_img_bigger = cv2.copyMakeBorder(cropped_img, 200, 200, 200, 200, cv2.BORDER_CONSTANT, value=(255, 255, 255))

                    extractedInformation = pytesseract.image_to_string(cropped_img_bigger).strip()

                    extracted.append(extractedInformation)
                
                new_row = pd.DataFrame({
                                    "NOMOR": [extracted[0]],
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
                                    "C.4 TANGGAL": [extracted[16]]
                                })
                df_all_data_extracted = pd.concat([df_all_data, new_row]).reset_index(drop=True)
                return(df_all_data_extracted)

            df_all_data_extracted = extract_text(image=img, coordinates=coordinates, all_data=df_all_data)

            df_all_data_extracted_combined = pd.concat([df_all_data_extracted_combined, df_all_data_extracted]).reset_index(drop=True)

            st.dataframe(df_all_data_extracted_combined)


        



        
    else:
        st.warning('You need to upload zip type file')
    
    
    # for pdf_path in glob.glob("/content/*.jpg")

else :
    st.error("You have to upload pdf folder in the sidebar")



# images = convert_from_path(f"D:/Data Science/Project etc/OCR pajak/tes ocr/tes ocr/M_01-DOC001_SPT_Unifikasi_English_BPU_ID-fo-xsl_DN2025173976628119321.pdf", 500)



# for i, image in enumerate(images):
#     fname = 'image'+str(i+2)+'.jpg'
#     image.save(fname, "JPEG")


        
        # tab1, tab2 = st.tabs(["Setting", "Run Apps"])
        # with tab1 :


#         col_1, col_2, col_3, col_4, col_5 = st.columns(5)
#         with st.container():
#             with col_1:
#                 user_input_npwp = 'IDENTITAS_PENERIMA_PENGHASILAN'
#         with st.container():
#             with col_2:
#                 user_input_perusahaan = 'NAMA_PENERIMA_PENGHASILAN'
#         with st.container():
#             with col_3:
#                 user_input_masa_pajak = 'MASA_PAJAK'
#         with st.container():
#             with col_4:
#                 user_input_tahun_pajak = 'TAHUN_PAJAK'
#         with st.container():
#             with col_5:
#                 user_input_ID = 'ID_SISTEM'

#         submit_button_clicked = st.button("Submit", type="primary", use_container_width=True)

#         if submit_button_clicked :


#             a = os.listdir(os.path.join(os.getcwd(),os.path.splitext(user_input_folder.name)[0]))
#             lst = []
#             for x in a :
#                 lst.append(os.path.splitext(x)[0][-36 :])

#             if len(lst) != len(df) :
#                 st.error('Data length not matched!')
#             else :

#                 for i in range(len(lst)):
#                     matching_index = df.index[df[user_input_ID] == lst[i]]
#                     nama_perusahaan = df.loc[matching_index, user_input_perusahaan]
#                     npwp_perusahaan = df.loc[matching_index, user_input_npwp]
#                     nama_npwp_perusahaan = str(nama_perusahaan.item()) + ' (' + str(npwp_perusahaan.item()) + ')'
#                     tahun_pajak = df.loc[matching_index, user_input_tahun_pajak]
#                     masa_pajak = df.loc[matching_index, user_input_masa_pajak]
#                     tahun_masa_pajak = str(tahun_pajak.item()) + '-' + str(masa_pajak.item())

#                     result_path = os.path.join(os.getcwd(),'Result')
#                     if os.path.exists(result_path) == False:
#                         os.mkdir(result_path)
#                     if os.path.exists(os.path.join(result_path, nama_npwp_perusahaan)) == False:
#                         os.mkdir(os.path.join(result_path, nama_npwp_perusahaan))
#                     if os.path.exists(os.path.join(result_path, nama_npwp_perusahaan, tahun_masa_pajak)) == False:
#                         os.mkdir(os.path.join(result_path, nama_npwp_perusahaan, tahun_masa_pajak))
#                     path_to_save = os.path.join(result_path, nama_npwp_perusahaan, tahun_masa_pajak)
                    
#                     shutil.copy(glob.glob(os.path.join(os.getcwd(),os.path.splitext(user_input_folder.name)[0],'*pdf'))[i], path_to_save)
                
#                 # st.write(os.listdir(os.getcwd()))

#                 shutil.make_archive('Result', 'zip', result_path)
                
#                 result_path_zipped = os.path.join(os.getcwd(),'Result.zip')

#                 with open(result_path_zipped, "rb") as fp :
#                     button_clicked = st.download_button(label=':cloud: Download Result', type="secondary", data=fp, file_name='Result.zip', mime="application/zip")




#     else :
#         st.error("You have to upload pdf folder in the sidebar")

# else :
#     st.error("You have to upload a csv or an excel file in the sidebar")






