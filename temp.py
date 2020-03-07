import sys
import time
import os
# import xmltodict
# import magicdate
from watchdog.observers import Observer
from watchdog.events import PatternMatchingEventHandler

# from .models import Media


class MyHandler(PatternMatchingEventHandler):
    patterns=["*.xml"]


    def process(self, event):
        """
        event.event_type
            'modified' | 'created' | 'moved' | 'deleted'
        event.is_directory
            True | False
        event.src_path
            path/to/observed/file
        """
        sdf=0
        # with open(event.src_path, 'r') as xml_source:
        #     pass
            # xml_source.write()
            # xml_string = xml_source.read()
            # parsed = xmltodict.parse(xml_string)
            # element = parsed.get('Pulsar', {}).get('OnAir', {}).get('media')
            # if not element:
            #     return
            #
            # media = Media(
            #     title=element.get('title1'),
            #     description=element.get('title3'),
            #     media_id=element.get('media_id1'),
            #     hour=magicdate(element.get('hour')),
            #     length=element.get('title4')
            # )
            # media.save()

    def on_modified(self, event):
        self.process(event)

    def on_created(self, event):
        self.process(event)

    def on_any_event(self, event):
        print(str(event))
        self.process(event)




if __name__ == '__main__':
    args = sys.argv[1:]
    observer = Observer()
    observer.schedule(MyHandler(), path=os.path.join("d:/test"), recursive= True)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()

    observer.join()
