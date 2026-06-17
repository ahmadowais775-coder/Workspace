# Workspace

A collection of projects including presentations, study notes, and Islamic education materials.

## Folder Structure

### Presentations/
AI Overview PowerPoint presentation covering the fundamentals of Artificial Intelligence.
- `create_ppt.py` - Python script to generate the presentation
- `AI_Overview_Presentation.pptx` - The generated 10-slide presentation

### Study_Notes/
Anatomy and physiology study notes and the scripts used to generate them.
- `create_joints_notes.py` - Script to generate joints and sutures notes
- `create_tissues_notes.py` - Script to generate tissues notes
- `Joints_and_Sutures_Notes.docx` - Generated study notes on joints and sutures
- `The_Tissues_Notes.docx` - Generated study notes on tissues

### Islamic_Majlis/
Islamic Majlis educational cards covering teachings that are often underestimated.
- `create_majlis_cards.py` - Script to generate the first set of cards
- `create_majlis_cards_part2.py` - Script to generate the second set of cards
- `Islamic_Majlis_Cards.html` - First set of Islamic Majlis cards
- `Islamic_Majlis_Cards_Part2.html` - Second set (27 lesser-known teachings)

## Requirements

To regenerate any of the documents, install the required dependencies:

```bash
pip install python-pptx python-docx
```

Then run the relevant script from within its folder, for example:

```bash
cd Presentations
python create_ppt.py
```
