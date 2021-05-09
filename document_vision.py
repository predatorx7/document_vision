#!/usr/bin/env python3

import enum
import os
from os.path import isfile, join
import time
import re
import multiprocessing
from sys import argv
from PIL import Image
import pytesseract
from pdf2image import convert_from_path
import docx
from pytesseract.pytesseract import Output

__cpu_count__ = multiprocessing.cpu_count()


def clamp(num: int, min_value: int, max_value: int) -> int:
    return max(min(num, max_value), min_value)


def __numberOfThreads__(percentageOfCpus=0.7) -> int:
    numberOfCpuToUse = round(__cpu_count__ * percentageOfCpus)

    # clamping to avoid comp issues
    return clamp(numberOfCpuToUse, 0, __cpu_count__)


thread_count = __numberOfThreads__()


def timedTask(callback):
    s = time.perf_counter()
    result = callback()
    elapsed = time.perf_counter() - s
    print(f"{callback} executed in {elapsed:0.2f} seconds.")
    return result

def valid_xml_char_ordinal(c):
    codepoint = ord(c)
    # conditions ordered by presumed frequency
    return (
        0x20 <= codepoint <= 0xD7FF or
        codepoint in (0x9, 0xA, 0xD) or
        0xE000 <= codepoint <= 0xFFFD or
        0x10000 <= codepoint <= 0x10FFFF
        )

def makeXMLCompatible(input_string):
    return ''.join(c for c in input_string if valid_xml_char_ordinal(c))


class OutputMode(enum.Enum):
    docx = 0
    txt = 1

output_file_name = None

class DocumentVision:
    def performDetectionOnImage(filename: str):
        image = Image.open(filename)
        return pytesseract.image_to_string(image, lang='eng')

    def detectImagesToTxt(imageFilePaths: list):
        pageNumber = 0
        folderPath = os.path.dirname(imageFilePaths[0])
        detectionResult = f'{output_file_name}.txt'

        print(f'Writing detection result in {detectionResult}\n')

        with open(detectionResult, 'w+') as f:
            for pageImagePath in imageFilePaths:
                print(f"performing detection for {pageImagePath}")

                def voidCallback():
                    return DocumentVision.performDetectionOnImage(pageImagePath)

                text = timedTask(voidCallback)

                f.write(text)

                print(f'done writing page {pageNumber}')

                f.write(f'\n\n\n')

                pageNumber += 1

    def detectImagesToDocx(imageFilePaths: list):
        pageNumber = 0
        detectionResult = f'{output_file_name}.docx'

        output_docx = docx.Document()

        print(f'Writing detection result in {detectionResult}\n')

        for pageImagePath in imageFilePaths:
            print(f"performing detection for {pageImagePath}")

            def voidCallback():
                return DocumentVision.performDetectionOnImage(pageImagePath)

            text = timedTask(voidCallback)

            output_docx.add_paragraph(makeXMLCompatible(text))
            output_docx.add_page_break()

            print(f'done writing page {pageNumber}')

            pageNumber += 1

        output_docx.save(detectionResult)

    def detect(filename: str, output_mode: OutputMode) -> list:
        assert(filename.endswith(".pdf"))

        print(f'Starting detection on file {filename}')

        print(f'Converting pdf to images with {thread_count} thread(s)')

        def pdfConversionTask():
            return convert_from_path(
                filename, 500, thread_count=thread_count)

        pdfpages = timedTask(pdfConversionTask)

        print(f'Converting from pdf to images is successful')

        pageImagesPath = []

        saveDir = f'{filename}.d'

        try:
            os.stat(saveDir)
        except:
            os.mkdir(saveDir)

        pageNumber = 0

        for page in pdfpages:
            pageSavePath = f'{saveDir}/{pageNumber}.jpg'
            print(f"Saving page {page} at {pageSavePath}")
            page.save(pageSavePath, 'JPEG')
            pageImagesPath.append(pageSavePath)
            pageNumber += 1

        if (output_mode is OutputMode.docx):
            DocumentVision.detectImagesToDocx(pageImagesPath)
        else:
            DocumentVision.detectImagesToTxt(pageImagesPath)


def tryint(s):
    try:
        return int(s)
    except:
        return s


def alphanum_key(x):
    """ Turn a string into a list of string and number chunks.
                    "z23a" -> ["z", 23, "a"]
    """
    s = os.path.basename(x)
    return [tryint(c) for c in re.split('([0-9]+)', s)]


def sort_nicely(l):
    """ Sort the given list in the way that humans expect.
    """
    l.sort(key=alphanum_key)


def detectFiles(folderPath: str, outputmode: OutputMode):
    imageFilePaths = [join(folderPath, f) for f in os.listdir(
        folderPath) if isfile(join(folderPath, f))]
    sort_nicely(imageFilePaths)
    if (outputmode is OutputMode.docx):
        DocumentVision.detectImagesToDocx(imageFilePaths)
    else:
        DocumentVision.detectImagesToTxt(imageFilePaths)

def main():
    option = argv[1]
    outputmode = None
    targetFilePath = None
    hasOptions = option.startswith('-')
    doPdfToJpg = True
    global output_file_name
    if (hasOptions):
        targetFilePath = os.path.abspath(argv[2])
        output_file_name = os.path.abspath(argv[3])
        if 'd' in option:
            print('output set to docx mode')
            outputmode = OutputMode.docx
        if 's' in option:
            print('Skipping pdf to images conversion')
            doPdfToJpg = False
    else:
        print('output set to txt mode')
        outputmode = OutputMode.txt
        targetFilePath = os.path.abspath(argv[1])
        output_file_name = os.path.abspath(argv[2])
    print(f'Target file(s): {targetFilePath}')
    if not doPdfToJpg:
        detectFiles(targetFilePath, outputmode)
    else:
        DocumentVision.detect(targetFilePath, outputmode)
    
        
if __name__ == "__main__":
    main()
