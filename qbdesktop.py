import win32com.client
import pandas as pd
import os
from xml.etree import ElementTree
from xml.etree.ElementTree import Element 
from dataclasses import dataclass 

#================================================================
#CONST
#================================================================
DEFAULT_COMPANY_FILE = f"C:\\Users\\Public\\Documents\\Intuit\\QuickBooks\\Company Files"
REQUESTS_PATH = os.path.join(os.getcwd(), "all_requests.xml")

#================================================================
# UTIL
#================================================================



#================================================================
# DATACLASSES
#================================================================

@dataclass(slots=True)
class FileMode:
    do_not_care = win32com.client.Dispatch("QBXMLRP2Lib.QBFileMode.qbFileOpenDoNotCare")
    single_user = win32com.client.Dispatch("QBXMLRP2Lib.QBFileMode.qbOpenSingleUser")
    multi_user = win32com.client.Dispatch("QBXMLRP2Lib.QBFileMode.qbOpenMultiUser")
    restore = win32com.client.Dispatch("QBXMLRP2Lib.QBFileMode.qbFileRestore")
    condense = win32com.client.Dispatch("QBXMLRP2Lib.QBFileMode.qbFileCondense")
    data_recovery = win32com.client.Dispatch("QBXMLRP2Lib.QBFileMode.qbFileAutoDataRecovery")

    
@dataclass(slots=True)
class QBAgingReport:
    BudgetSummary = 0 
    CustomDetail = 1 
    CustomSummary = 2
    GeneralDetail = 3 
    GeneralSummary = 4
    Job = 5
    PayrollDetail = 6
    PayrollSummary =7
    Time = 8
    
    
@dataclass(slots=True)
class QBJobReport:
    ItemEstimatesVsActuals = 0
    
@dataclass(slots=True)
class QBAgingRequest:
    BudgetSummary = "BudgetSummaryReportQueryRq"
    CustomDetail = "CustomDetailReportQueryRq"
    CustomSummary = "CustomSummaryReportQueryRq"
    GeneralDetail = "GeneralDetailReportQueryRq"
    GeneralSummary = "GeneralSummaryReportQueryRq"
    Job = "JobReportQueryRq"
    PayrollDetail = "PayrollDetailReportQueryRq"
    PayrollSummary = "PayrollSummaryReportQueryRq"
    Time = "	TimeReportQueryRq"
    

#================================================================
# TYPES
#================================================================


class qbXML:
    pass


class Param:
    def __init__(self, name: str, value: str) -> None:
        self.name = name
        self.value = value
    
    def read(self): 
        return f'{self.name}="{self.value}"'
        

class Element(qbXML):
    def __init__(self, name: str, value: str, indent: int=3) -> None:
        self.name = name
        self.value = value
        self.indent = indent
        self._indent = "\t"*indent
        self.statement = f"{self._indent}<{self.name}{self.value}</{self.name}>"
        super().__init__()
        
    def read(self) -> str:
        return self.statement
    

class Aggregate(qbXML):
    def __init__(self, name: str, elements: list[Element]=[], indent: int=2, params=[Param]) -> None:
        super().__init__()
        self.params: list[Param] = params 
        self.name = name
        self.elements = elements
        self.indent = indent
        
        self.opening = f"<{self.name}"
        for param in self.params:
            self.opening += " "
            self.opening+=param.read()
        self.opening+=" >"
        self.closing = f"</{self.name}>"
        self.objects = [self.opening, self.closing]

    def add_element(self, element: Element) -> None:
        self.elements.append(element)
        
    def read(self) -> str:
        objs = self.objects
        for element in self.elements:
            objs.insert(1, element.read())
        return "\n".join(objs)
        
        
class MessageAggregate(qbXML):
    def __init__(self, name: str, aggregates: list[Aggregate]=[], indent: int=1, params:list[Param] =[]) -> None:
        super().__init__()
        self.params: list[Param] = params 
        self.name = name
        self.aggregates = aggregates
        self.indent = indent
        self.opening = f"<{self.name}"
        for param in self.params:
            self.opening += " "
            self.opening+=param.read()
        self.opening+=" >"
        self.closing = f"</{self.name}>"
        self.objects = [self.opening, self.closing]
    
    def add_aggregate(self, aggregate: Aggregate) -> None:
        self.aggregates.append(aggregate)
        
    def read(self) -> str:
        objs = self.objects
        for agg in self.aggregates:
            objs.insert(1, agg.read())
        return "\n".join(objs)


#================================================================
# CLASSES
#================================================================

class QuickBooksResponse:
    def __init__(self, request: str=None, response: str=None) -> None:
        self.request = request
        self.response = response
        self._path = os.path.join(os.getcwd(), "temp1.xml")
        with open(self._path, 'r') as xml:
            xml.write(self.response)
        self.dataframe = pd.read_xml(self._path)
        os.remove(self._path)
    
    def save_as_excel(self, filepath: str):
        pass
        


class RequestProcessor:
    def __init__(self, app_name: str) -> None:
        try:
            self.qbxmlrp = win32com.client.Dispatch("QBXMLRP2.RequestProcessor")
        except Exception as e:
            print(e)
            exit()
        self.app_name = app_name
        self.ticket = None

    def __enter__(self):
        self.open_connection()
        self.begin_session()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.end_session()
        self.close_connection()

    def open_connection(self):
        self.qbxmlrp.OpenConnection("", self.app_name)

    def close_connection(self):
        self.qbxmlrp.CloseConnection()

    def begin_session(self):
        self.ticket = self.qbxmlrp.BeginSession("", 0)

    def end_session(self):
        self.qbxmlrp.EndSession(self.ticket)
        self.ticket = None

    def process_request(self, request):
        return self.qbxmlrp.ProcessRequest(self.ticket, request)
    
    

class SessionManager:
    def __init__(self, app_id, app_name, company_file_path):
        self.qb_session_manager = win32com.client.Dispatch("QBFC13.QBSessionManager")
        self.app_id = app_id
        self.app_name = app_name
        self.company_file_path = company_file_path
        self.qb_session = None

    def __enter__(self):
        self.begin_session()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.end_session()

    def begin(self):
        """used to manage the session with QuickBooks.
        """
        self.qb_session_manager.OpenConnection('', self.app_name)
        self.qb_session_manager.BeginSession(self.company_file_path, self.app_id)

        self.qb_session = self.qb_session_manager.Session
        self.qb_session.OpenEntity('Customer', '')

    def end(self):
        """used to manage the session with QuickBooks.
        """
        if self.qb_session:
            self.qb_session.Close()
            self.qb_session = None

        self.qb_session_manager.EndSession()
        self.qb_session_manager.CloseConnection()

    def create_customer(self, customer_data):
        """create a new customer in QuickBooks. 
        It takes customer data in XML format as an argument, 
        sends the customer add request to QuickBooks, 
        and returns the customer details.

        Args:
            customer_data (_type_): _description_

        Raises:
            Exception: _description_

        Returns:
            _type_: _description_
        """
        customer_add_rq = self.qb_session.CreateMsgSetRequest('US', '2.0').AppendCustomerAddRq()
        customer_add_rq.CustomerAdd.FromXmlString(customer_data)
        response = self.qb_sessionManager.DoRequests(customer_add_rq).ResponseList.GetAt(0)

        if response.StatusCode != 0:
            raise Exception(response.StatusMessage)

        return response.Detail

    def query_customers(self):
        """ query customers in QuickBooks. 
        It sends a customer query request to QuickBooks and returns the customer details

        Raises:
            Exception: _description_

        Returns:
            _type_: _description_
        """
        customer_query_rq = self.qb_session.CreateMsgSetRequest('US', '2.0').AppendCustomerQueryRq()
        response = self.qb_sessionManager.DoRequests(customer_query_rq).ResponseList.GetAt(0)

        if response.StatusCode != 0:
            raise Exception(response.StatusMessage)

        return response.Detail
    

class RequestProcessorDialog:
    def __init__(self, app_id, app_name, company_file_path):
        """ simple interface for showing the QuickBooks request processor 
        dialog and retrieving the response using the QBXMLRP2UI.RequestProcessorDialog
        COM object. The constructor takes the QuickBooks Application ID, Application Name, 
        and Company File Path as arguments.

        Args:
            app_id (_type_): _description_
            app_name (_type_): _description_
            company_file_path (_type_): _description_
        """
        self.qb_request_processor = win32com.client.Dispatch("QBXMLRP2UI.RequestProcessorDialog")
        self.app_id = app_id
        self.app_name = app_name
        self.company_file_path = company_file_path

    def show(self, xml_request):
        """display the request processor dialog with the specified XML request. 

        Args:
            xml_request (_type_): _description_
        """
        self.qb_request_processor.Reset()
        self.qb_request_processor.Init(self.app_id, self.app_name, self.company_file_path)
        self.qb_request_processor.Show(xml_request)

    def is_done(self):
        """used to check if the dialog has completed processing the 
        request and returned a response. It returns True if the dialog
        is done processing the request and False otherwise.
        """
        return self.qb_request_processor.IsDone()

    def get_response(self):
        """ retrieve the response from the dialog once it has completed processing the request. 
        It returns the response as a string.

        Returns:
            _type_: _description_
        """
        return self.qb_request_processor.Response

    def close(self):
        self.qb_request_processor.EndSession()


class WebConnector:
    def __init__(self, url):
        """simple interface for interacting with the QuickBooks Web Connector 
        using the QBWebConnector.QBWebConnectorSvc COM object. 
        The constructor takes the URL of the QBWC web service as an argument.

        Args:
            url (_type_): _description_
        """
        self.qb_web_connector = win32com.client.Dispatch("QBWebConnector.QBWebConnectorSvc")
        self.url = url

    def get_version(self):
        """ used to retrieve the version

        Returns:
            _type_: _description_
        """
        return self.qb_web_connector.get_Version()

    def get_error(self):
        """ used to retrieve the errors

        Returns:
            _type_: _description_
        """
        return self.qb_web_connector.get_Error()

    def get_ticket(self):
        """ used to retrieve the ticket

        Returns:
            _type_: _description_
        """        
        return self.qb_web_connector.get_Ticket()

    def close(self, ticket):
        """ used to close the connection with QuickBooks using the specified ticket number.

        Returns:
            _type_: _description_
        """
        return self.qb_web_connector.closeConnection(ticket)

    def process_request(self, ticket, str_request):
        """send a request to QuickBooks using the specified ticket 
        number and request XML string. The method returns a string 
        containing the response from QuickBooks.

        Args:
            ticket (_type_): _description_
            str_request (_type_): _description_

        Returns:
            _type_: _description_
        """
        return self.qb_web_connector.processRequest(ticket, str_request, self.url)

    def receive_response(self, ticket, response, hresult, message):
        """ receive a response from QuickBooks using the specified ticket number, 
        response XML string, HRESULT, and message. 
        The method returns True if the response was received successfully and False otherwise.

        Args:
            ticket (_type_): _description_
            response (_type_): _description_
            hresult (_type_): _description_
            message (_type_): _description_

        Returns:
            _type_: _description_
        """
        return self.qb_web_connector.receiveResponse(ticket, response, hresult, message, self.url)

    def send_request(self, ticket, company_file_path, qb_file_mode, request):
        """ send a request to QuickBooks using the specified ticket number,
        company file path, QB file mode, and request XML string. The method 
        returns True if the request was sent successfully and False otherwise.

        Args:
            ticket (_type_): _description_
            company_file_path (_type_): _description_
            qb_file_mode (_type_): _description_
            request (_type_): _description_

        Returns:
            _type_: _description_
        """
        return self.qb_web_connector.sendRequest(ticket, company_file_path, qb_file_mode, request, self.url)



class RequestAllData:
    
    def __init__(self, company_file_path, qb_file_mode):
        """sends a request for all data in the QuickBooks 
        company file specified by company_file_path and 
        returns the response as a QBXMLRP2.ResponseReader object

        Args:
            company_file_path (_type_): _description_
            qb_file_mode (_type_): _description_
        """
        self.qbxmlrp = win32com.client.Dispatch("QBXMLRP2.RequestProcessor")
        self.qbxmlrp.OpenConnection('', 'Python QuickBooks Connector')
        self.qbxmlrp.BeginSession(company_file_path, qb_file_mode)

    def __del__(self):
        self.qbxmlrp.EndSession()
        self.qbxmlrp.CloseConnection()

    def send_request(self):
        request = self.qbxmlrp.CreateMsgSetRequest('US', '13.0')
        request.AppendRequestForFullSync()
        response_str = self.qbxmlrp.ProcessRequest(request.ToXMLString(), '')
        response = win32com.client.Dispatch("QBXMLRP2.ResponseReader")
        response.LoadString(response_str)
        return response
