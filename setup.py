from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="email_batch_tool",
    version="0.1.0",
    author="Your Name",
    author_email="your.email@example.com",
    description="简要描述您的项目",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/yourusername/email_batch_tool",
    packages=find_packages(),
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.6",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
    ],
    python_requires='>=3.6',
    install_requires=[
        "msal>=1.20.0",
        "requests>=2.25.1",
        "beautifulsoup4>=4.12.0",
        "html5lib>=1.1",
    ],
    extras_require={
        "dev": [
            "pytest>=6.0",
            "pytest-cov>=2.0",
            "flake8>=3.9",
            "black>=21.0",
        ],
    },
    entry_points={
        'console_scripts': [
            'email_batch_tool=email_batch_tool.main:main',
        ],
    },
)
