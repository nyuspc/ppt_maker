import setuptools

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setuptools.setup(
    name="ppt_maker",
    version="0.0.1",
    author="Shuo Sun, Pengcheng Song",
    author_email="smth_spc@hotmail.com",
    description="Make PowerPoint slides with template and data",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/nyuspc/ppt_maker",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python",
        "Intended Audience :: Financial and Insurance Industry",
        "Topic :: Multimedia :: Graphics",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
)
