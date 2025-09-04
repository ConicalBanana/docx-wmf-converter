# docx-wmf-converter
A tool for converting embedded WMF/EMF file to SVG file inside DOCX.

## Motivation

Chinese version: [zhihu-article](https://zhuanlan.zhihu.com/p/1946891527935726678)

When converting DOCX to PDF in non-Windows platform, [pandoc](https://pandoc.org) will raise error:

```
Error producing PDF.

! LaTeX Error: Cannot determine size of graphic in path_to_your_embedded_emf.emf (no Bou
ndingBox).
```

The fundamental issue is that `xelatex` or `pdflatex` could not render `EMF`/`WMF` correctly (at least I cannot find a perfect solution). So this simple script would convert all the `EMF`/`WMF` images to `SVG` format and the output document could be converted to `PDF` format by [pandoc](https://pandoc.org)


## Dependencies

### Executable

1. libwmf [github-repo](https://github.com/caolanm/libwmf)
2. libemf2svg [github-repo](https://github.com/kakwa/libemf2svg)

### Python

3. BeautifulSoup [official-website](https://www.crummy.com/software/BeautifulSoup/)

## Installation

### MacOS

```bash
brew install libwmf
brew install libemf2svg

# Your conda environment settings
# ...
pip install bs4
```

### Linux

To be completed...

### Acknowledgements

Thanks for the wonderful libraries `libwmf` and `linemf2svg` :)