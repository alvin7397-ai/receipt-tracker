import os, json, base64, sqlite3, io
from datetime import datetime
from flask import Flask, request, jsonify, render_template, send_file, session, redirect, url_for
import anthropic
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter