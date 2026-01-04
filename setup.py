"""
Setup script per Code Comparator
Applicazione desktop per confrontare liste basate su colonne chiave

Copyright © 2025 Biblioteca Morante
Licenza: MIT License
"""

from setuptools import setup
import os


def read_file(filename):
    """Legge il contenuto di un file (es. README.md)"""
    here = os.path.abspath(os.path.dirname(__file__))
    filepath = os.path.join(here, filename)
    
    if os.path.exists(filepath):
        with open(filepath, encoding='utf-8') as f:
            return f.read()
    return ""


setup(
    # ============================================================
    # INFORMAZIONI PROGETTO
    # ============================================================
    name='code-comparator',
    version='1.0.0',
    
    author='Biblioteca Morante',
    author_email='bibliotecamorante@gmail.com',
    
    description='Applicazione desktop per confrontare liste basate su colonne chiave (ISBN, inventario, ecc.)',
    long_description=read_file('README.md'),
    long_description_content_type='text/markdown',
    
    url='https://github.com/bibliotecamorante/code-comparator',
    license='MIT',
    
    # ============================================================
    # CLASSIFICATORI PYPI
    # ============================================================
    classifiers=[
        # Stato sviluppo
        'Development Status :: 4 - Beta',
        
        # Pubblico target
        'Intended Audience :: End Users/Desktop',
        'Intended Audience :: Education',
        'Intended Audience :: Information Technology',
        
        # Categoria
        'Topic :: Office/Business',
        'Topic :: Database',
        'Topic :: Utilities',
        
        # Licenza
        'License :: OSI Approved :: MIT License',
        
        # Versioni Python supportate
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.9',
        'Programming Language :: Python :: 3.10',
        'Programming Language :: Python :: 3.11',
        'Programming Language :: Python :: 3.12',
        'Programming Language :: Python :: 3.13',
        
        # Sistema operativo
        'Operating System :: OS Independent',
        'Operating System :: Microsoft :: Windows',
        'Operating System :: MacOS',
        'Operating System :: POSIX :: Linux',
        
        # Lingua
        'Natural Language :: Italian',
        'Natural Language :: English',
        
        # Ambiente
        'Environment :: X11 Applications :: Qt',
    ],
    
    # ============================================================
    # KEYWORDS
    # ============================================================
    keywords='comparison excel data library management pyqt6 isbn inventory catalog',
    
    # ============================================================
    # MODULI DA INSTALLARE
    # ============================================================
    # ⚠️ IMPORTANTE: Il file principale si chiama code_comparator.py
    py_modules=['code_comparator'],
    
    # ============================================================
    # DIPENDENZE
    # ============================================================
    install_requires=[
        'PyQt6>=6.4.0',      # Interfaccia grafica
        'pandas>=2.0.0',     # Elaborazione dati
        'openpyxl>=3.1.0',   # Export Excel formattato
    ],
    
    # Dipendenze opzionali (installabili con: pip install code-comparator[dev])
    extras_require={
        'dev': [
            'pyinstaller>=5.0',  # Per creare eseguibili
            'pytest>=7.0',       # Per test (se li aggiungi)
        ],
    },
    
    # ============================================================
    # VERSIONE PYTHON RICHIESTA
    # ============================================================
    python_requires='>=3.9',
    
    # ============================================================
    # ENTRY POINTS (come lanciare l'app da terminale)
    # ============================================================
    entry_points={
        'gui_scripts': [
            # Comando: code-comparator
            # Esegue: code_comparator.py → funzione main()
            'code-comparator=code_comparator:main',
        ],
    },
    
    # ============================================================
    # FILE DA INCLUDERE NEL PACCHETTO
    # ============================================================
    include_package_data=True,  # Usa MANIFEST.in
    
    # ============================================================
    # LINK PROGETTO
    # ============================================================
    project_urls={
        'Homepage': 'https://github.com/bibliotecamorante/code-comparator',
        'Bug Reports': 'https://github.com/bibliotecamorante/code-comparator/issues',
        'Source': 'https://github.com/bibliotecamorante/code-comparator',
        'Documentation': 'https://github.com/bibliotecamorante/code-comparator#readme',
    },
    
    # ============================================================
    # METADATI AGGIUNTIVI
    # ============================================================
    platforms=['any'],
    zip_safe=False,  # Non comprimere il pacchetto (consigliato per GUI)
)
