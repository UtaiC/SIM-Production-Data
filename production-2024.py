import streamlit as st
import pandas as pd
from PIL import Image
##################################
# Css Style #####################
with open('style.css') as modi:
    css = f'<style>{modi.read()} </style>'
    st.markdown(css, unsafe_allow_html=True)
# Banner #################################
banner_image = Image.open('Banner-Prod.jpg')
st.image(banner_image, width=800)
##########################################
st.header('SIM Production Page')