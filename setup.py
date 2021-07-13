setup(
    name="AlopexMM-Email-Sender-For-Applysys",
    version="1.0.0",
    description="Send emails from applysys excel",
    long_description=README,
    long_description_content_type="text/markdown",
    url="",
    author="AlopexMM",
    author_email="mariomori00@gmail.com",
    license="MIT",
    classifiers=[
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
    ],
    packages=["src"],
    include_package_data=True,
    install_requires=["xlrd","pyqt5","pyqt5-tools"],
    entry_points={"console_scripts":["alopexmm=src.__main__:main"]},
)
