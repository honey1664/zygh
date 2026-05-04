# chat_floating_v2.py
import streamlit as st

def render_floating_chat():
    """渲染右下角悬浮聊天图标 - 点击在新标签页打开独立聊天页面"""

    # 添加右下角悬浮图标的CSS和HTML
    st.markdown("""
    <style>
    /* 悬浮按钮样式 */
    .floating-chat-btn {
        position: fixed;
        bottom: 30px;
        right: 30px;
        width: 65px;
        height: 65px;
        border-radius: 50%;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        display: flex;
        align-items: center;
        justify-content: center;
        cursor: pointer;
        box-shadow: 0 4px 20px rgba(102, 126, 234, 0.4);
        transition: all 0.3s ease;
        z-index: 999;
        font-size: 30px;
        text-decoration: none;
        animation: pulse 2s infinite;
        border: none;
    }

    .floating-chat-btn:hover {
        transform: scale(1.1);
        box-shadow: 0 6px 25px rgba(102, 126, 234, 0.5);
        text-decoration: none;
    }

    @keyframes pulse {
        0% {
            box-shadow: 0 0 0 0 rgba(102, 126, 234, 0.4);
        }
        70% {
            box-shadow: 0 0 0 15px rgba(102, 126, 234, 0);
        }
        100% {
            box-shadow: 0 0 0 0 rgba(102, 126, 234, 0);
        }
    }
    </style>

    <a href="chat_page" target="_blank" class="floating-chat-btn" rel="noopener noreferrer">
        💬
    </a>
    """, unsafe_allow_html=True)