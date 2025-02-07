import streamlit as st
from PIL import Image
import requests
from io import BytesIO

image_url = "https://ideaconsulteu-my.sharepoint.com/:f:/r/personal/alexanderl_ideaconsult_be/Documents/Apps/Microsoft%20Forms/Rapportage%20strategisch%20project%20(met%20foto%27s)?csf=1&web=1"


response = requests.get(image_url)

if response.status_code == 200:
    image = Image.open(BytesIO(response.content))
    st.image(image, caption="Loaded Image", use_column_width=True)
else:
    st.error("Could not load image. Check if the link is publicly accessible.")
