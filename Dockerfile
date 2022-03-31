FROM python:3.8
ADD bio_txt_to_xls.py .
RUN pip install pandas XlsxWriter 
ENTRYPOINT ["python", "./bio_txt_to_xls.py"]