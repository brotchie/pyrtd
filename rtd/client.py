import pythoncom

from win32com import client, universal
from win32com.server.util import wrap

# Thanks to Chris Nilsson for these constants.
# Typelib info for Excel 2007.
EXCEL_TLB_GUID = '{00020813-0000-0000-C000-000000000046}'
EXCEL_TLB_LCID = 0
EXCEL_TLB_MAJOR = 1
EXCEL_TLB_MINOR = 6

# Register the two RTD interfaces defined in the Excel typelib.
universal.RegisterInterfaces(EXCEL_TLB_GUID, 
                             EXCEL_TLB_LCID, EXCEL_TLB_MAJOR, EXCEL_TLB_MINOR,
                             ['IRtdServer','IRTDUpdateEvent'])
MAX_REGISTERED_TOPICS = 1024

class RTDClient(object):
    """
    Implements a Real-Time-Data (RTD) client for accessing
    COM datasources that provide an IRtdServer interface.

     - Implements the IRTDUpdateEvent interface and if used
       in event driven mode only calls RefreshData when
       new data is available.

    """

    _com_interfaces_ = ['IRTDUpdateEvent']
    _public_methods_ = ['Disconnect', 'UpdateNotify']
    _public_attrs_ = ['HeartbeatInterval']

    def __init__(self, classid):
        self._classid = classid
        self._rtd = None

        self._data_ready = False

        self._topic_to_id = {}
        self._id_to_topic = {}
        self._topic_values = {}
        self._last_topic_id = 0

    def connect(self, event_driven=True):
        """
        Connects to the RTD server.
        
        Set event_driven to false if you to disable update notifications.
        In this case you'll need to call refresh_data manually.

        """

        dispatch = client.Dispatch(self._classid)
        self._rtd = client.CastTo(dispatch, 'IRtdServer')
        if event_driven:
            self._rtd.ServerStart(wrap(self))
        else:
            self._rtd.ServerStart(None)

    def update(self):
        """
        Check if there is data waiting and call RefreshData if
        necessary. Returns True if new data has been received.

        Note that you should call this following a call to
        pythoncom.PumpWaitingMessages(). If you neglect to
        pump the message loop you'll never receive UpdateNotify
        callbacks.

        """
        if self._data_ready:
            self._data_ready = False
            self.refresh_data()
            return True
        else:
            return False

    def refresh_data(self):
        """
        Grabs new data from the RTD server.

        """

        (ids, values), count = self._rtd.RefreshData(MAX_REGISTERED_TOPICS)
        for id, value in zip(ids, values):
            assert id in self._id_to_topic
            topic = self._id_to_topic[id]
            self._topic_values[topic] = value

    def get(self, topic):
        """
        Gets the value of a registered topic. Returns None
        if no value is available. Throws an exception if
        the topic isn't registered.

        """

        assert topic in self._topic_to_id, 'Topic %s not registered.' % (topic,)
        return self._topic_values.get(topic)

    def register_topic(self, topic):
        """
        Registers a topic with the RTD server. The topic's value
        will be updated in subsequent data refreshes.

        """

        if topic not in self._topic_to_id:
            id = self._last_topic_id
            self._last_topic_id += 1
            
            self._topic_to_id[topic] = id
            self._id_to_topic[id] = topic

            self._rtd.ConnectData(id, (topic,), True)

    # Implementation of IRTDUpdateEvent.
    HeartbeatInterval = -1

    def UpdateNotify(self):
        self._data_ready = True

    def Disconnect(self):
        pass
