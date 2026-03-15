"""Post-process the DOCX to make it Word-compatible:
1. Remove w:background element
2. Remove displayBackgroundShape from settings
3. Remove the anchor hack from header (keep banner as inline image)
"""
import zipfile, os, re, shutil, sys

src = sys.argv[1] if len(sys.argv) > 1 else r"C:\git\y2k-portfolio\Y2K-AI-Portfolio.docx"
tmp = src + ".tmp"

with zipfile.ZipFile(src, 'r') as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
    for item in zin.namelist():
        data = zin.read(item)
        if item == 'word/document.xml':
            xml = data.decode('utf-8')
            # Remove w:background element
            xml = re.sub(r'<w:background[^/]*/>', '', xml)
            xml = re.sub(r'<w:background[^>]*>.*?</w:background>', '', xml, flags=re.DOTALL)
            data = xml.encode('utf-8')
        elif item == 'word/settings.xml':
            xml = data.decode('utf-8')
            xml = re.sub(r'<w:displayBackgroundShape[^/]*/>', '', xml)
            data = xml.encode('utf-8')
        elif 'header' in item and item.endswith('.xml'):
            xml = data.decode('utf-8')
            # Convert anchor back to inline if present
            xml = xml.replace('wp:anchor', 'wp:inline')
            # Remove anchor-specific attributes
            for attr in ['behindDoc', 'locked', 'layoutInCell', 'allowOverlap', 'simplePos', 'relativeHeight']:
                xml = re.sub(rf'\s+{attr}="[^"]*"', '', xml)
            # Remove positionH/positionV elements
            xml = re.sub(r'<wp:positionH[^>]*>.*?</wp:positionH>', '', xml, flags=re.DOTALL)
            xml = re.sub(r'<wp:positionV[^>]*>.*?</wp:positionV>', '', xml, flags=re.DOTALL)
            data = xml.encode('utf-8')
        zout.writestr(item, data)

shutil.move(tmp, src)
print("Fixed for Word compatibility")
