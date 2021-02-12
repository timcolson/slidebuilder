#%%
from pptx import Presentation

filename = "source1"
print(f"Read in powerpoint: {filename}.pptx")

#%%
ppt = Presentation("../" + filename + ".pptx")

print(ppt)
# %%
for slide in ppt.slides:
    print(f"Slide: {slide} - {slide.name}")
# %%

ppt.slides[1].name="TC-NewID"

ppt.save("../target/" + filename + "-pyout.pptx")

