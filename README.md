This repository features a curated collection of visually impressive presentations and brainstorming samples. 

**Workflow:**
* **Initial Generation:** Created using Gemini Notebook.
  * Base prompt use to specify the layout: *Use dark background with impactful, stunning imagery to complement the content, avoiding an artificial or AI-generated look & feel.*
* **Refinement:** Manually processed and polished before uploading.

**Contents:**
* `.pptx` files provided as read-only slide decks (and most try complete with speaker notes).
* Corresponding `.png` thumbnail images sharing the same filename, and deck with `_GO` (graphics only) are version without text elements.
* Decks can optimized and compressed using compress_pptx.py to do the compression prior upload.
* To resuse the image, you can use [clipdrop](https://clipdrop.co/) to remove the text. 
* For editable deck, create a `.pdf` version (or individual image) and convert the `.pdf`/image to editable `.pptx` using tools like [2pptx](https://2pptx.com/), [DeckEdit](https://deckedit.com/), [CopySlides](https://copyslides.com/), [Image2PPT](https://image2ppt.com/), [SlidesPilot](https://www.slidespilot.com/image-to-ppt), ...

**About the utilities:**
* `compress_pptx.py` compresses raster images inside a .pptx file while preserving OOXML relationships and content types. By default, it converts PNGs (and other raster images) to JPEGs, downscales images larger than --max-dim (default 2000px), and saves the output as \<name\>_\<yymm\>.pptx. Use --keep-png to retain PNG files and their transparency instead of converting them to JPG.\
Syntax:\
python compress_pptx.py slides.pptx\
python compress_pptx.py slides.pptx --keep-png --max-dim 1600 --quality 85 -o out/custom_deck.pptx
* `pptx_catalog.py` creates a single png which list the first 16 slides in thumbnail format.\
Syntax: python pptx_catalog.py slides.pptx
* `update_pptx_layout.py` updates the master template background to black, create a blank template for the slides and remove unused slide templates.\
Syntax: python update_pptx_layout.py slides.pptx
* `import_speaker_note.py` imports speaker note in word format into the corresponding powerpoint file; use `Slide #` as seperator for individual slide speaker note.\
Syntax: python import_speaker_note.py slides.pptx
