#!/bin/bash

jupyter nbconvert 'Reformat Greater MSP data.ipynb' --to slides --post serve --SlidesExporter.reveal_scroll=True # --SlidesExporter.reveal_theme=serif

