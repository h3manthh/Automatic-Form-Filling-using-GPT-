import os
import streamlit as st
from st_pages import Page, add_page_title, show_pages, hide_pages

def template_init():
    # Construct paths for pages
    home_page_path = os.path.join(os.path.dirname(__file__), '..', 'app.py')
    sign_page_path = os.path.join(os.path.dirname(__file__), '..', 'Pages', '1_Sign.py')
    
    # Show pages
    show_pages([
        Page(home_page_path, "Home", "🏠"),
        Page(sign_page_path, "Sign", "📜"),
    ])
    
    add_page_title()  # Optional method to add title and icon to current page

def template_sidebar():
    st.sidebar.markdown(
        """ <style> [data-testid='stSidebarNav'] > ul { min-height: 55vh; } </style> """,
        unsafe_allow_html=True,
    )
    st.markdown(
        """
        <style>
            .block-container {
                max-width: 100%;
            }
        </style>
    """,
        unsafe_allow_html=True,
    )

def confidential():
    from View.base import is_safe

    if is_safe():
        hide_pages(["sign"])
    else:
        from streamlit_extras.switch_page_button import switch_page

        switch_page("sign")
