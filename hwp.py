import olefile
import zlib
import struct
import sys

from io import BytesIO
from datetime import datetime
from datetime import timedelta
from pprint import pprint

class hwp_parser():
    def __init__(self, filename):
        self.filename = filename
        self.ole = olefile.OleFileIO(filename)
        self.ole_dir = ["/".join(i) for i in self.ole.listdir()]
        ## https://github.com/mete0r/pyhwp/blob/82aa03eb3afe450eeb73714f2222765753ceaa6c/pyhwp/hwp5/msoleprops.py#L151
        self.SUMMARY_INFORMATION_PROPERTIES = [
            dict(id=0x02, name='PIDSI_TITLE', title='Title'),
            dict(id=0x03, name='PIDSI_SUBJECT', title='Subject'),
            dict(id=0x04, name='PIDSI_AUTHOR', title='Author'),
            dict(id=0x05, name='PIDSI_KEYWORDS', title='Keywords'),
            dict(id=0x06, name='PIDSI_COMMENTS', title='Comments'),
            dict(id=0x07, name='PIDSI_TEMPLATE', title='Templates'),
            dict(id=0x08, name='PIDSI_LASTAUTHOR', title='Last Saved By'),
            dict(id=0x09, name='PIDSI_REVNUMBER', title='Revision Number'),
            dict(id=0x0a, name='PIDSI_EDITTIME', title='Total Editing Time'),
            dict(id=0x0b, name='PIDSI_LASTPRINTED', title='Last Printed'),
            dict(id=0x0c, name='PIDSI_CREATE_DTM', title='Create Time/Data'),
            dict(id=0x0d, name='PIDSI_LASTSAVE_DTM', title='Last saved Time/Data'),
            dict(id=0x0e, name='PIDSI_PAGECOUNT', title='Number of Pages'),
            dict(id=0x0f, name='PIDSI_WORDCOUNT', title='Number of Words'),
            dict(id=0x10, name='PIDSI_CHARCOUNT', title='Number of Characters'),
            dict(id=0x11, name='PIDSI_THUMBNAIL', title='Thumbnail'),
            dict(id=0x12, name='PIDSI_APPNAME', title='Name of Creating Application'),
            dict(id=0x13, name='PIDSI_SECURITY', title='Security'),
        ]

    def extract_data(self, name):
        stream = self.ole.openstream(name)
        data = stream.read()
        if any(i in name for i in ("BinData", "BodyText", "Scripts", "DocInfo")):
            return zlib.decompress(data,-15)
        else:
            return data

    def FILETIME_to_datetime(self, value):
        return datetime(1601, 1, 1, 0, 0, 0) + timedelta(microseconds=value / 10)

    def HwpSummaryInformation(self, data):
        info_data = []
        property_data = []
        return_data = []

        start_offset = 0x2c
        data_size_offset = struct.unpack("<L",data[start_offset:start_offset+4])[0]
        data_size = struct.unpack("<L",data[data_size_offset:data_size_offset+4])[0]
        property_count = struct.unpack("<L",data[data_size_offset+4:data_size_offset+8])[0]

        start_offset = data_size_offset + 8
        
        for i in range(property_count):
            property_ID = struct.unpack("<L",data[start_offset:start_offset+4])[0]
            unknown_data = struct.unpack("<L",data[start_offset+4:start_offset+8])[0]
            property_data.append({"property_ID":property_ID, "unknown_data":unknown_data})
            start_offset = start_offset + 8

        data = data[start_offset:]
        
        start_offset = 0x0
        for i in range(property_count):
            if data[start_offset:start_offset+4] == b"\x1f\x00\x00\x00":
                size = struct.unpack("<L",data[start_offset+4:start_offset+8])[0] * 2
                result = data[start_offset+8:start_offset+8+size]
                info_data.append(result.decode("utf-16-le"))

                start_offset = start_offset + 8 + size
                if data[start_offset:start_offset+2] == b"\x00\x00":
                    start_offset += 2

            elif data[start_offset:start_offset+4] == b"\x40\x00\x00\x00":
                date = struct.unpack("<Q", data[start_offset+4:start_offset+12])[0]
                start_offset = start_offset + 12
                info_data.append(self.FILETIME_to_datetime(date))

        for i in range(len(info_data)):
            for information in self.SUMMARY_INFORMATION_PROPERTIES:
                if information['id'] == property_data[i]['property_ID']:
                    return_data.append({"property_ID":property_data[i]['property_ID'], 
                                        "title":information['title'], 
                                        "name":information['name'], 
                                        "data":info_data[i],
                                        "unknown_data":property_data[i]['unknown_data']})
                    continue

        return return_data

    def run(self):
        #print("[*] ole dir : {}\n".format(self.ole_dir))
        for name in self.ole_dir:
            if "hwpsummaryinformation" in name.lower():
                data = self.extract_data(name)
                result = self.HwpSummaryInformation(data)
                pprint(result)

            if ".ps" in name.lower() or ".eps" in name.lower():
                print("[*] Extract eps file : {}".format(name.replace("/","_")))
                data = self.extract_data(name)
                f = open(name.replace("/","_"),"wb")
                f.write(data)
                f.close()
            
if __name__ == '__main__':
    try:
        a = hwp_parser(sys.argv[1])
        a.run()
    except OSError:
        print("[*] OSError !!")
    except TypeError:
        print("[*] TypeError !!")
