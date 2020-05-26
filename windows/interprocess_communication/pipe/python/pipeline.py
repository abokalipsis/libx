"""
Named pipelnie module
"""


import win32api
import win32pipe
import win32file
import threading
import pywintypes
import winerror
import time


class PipelineReader:
    def __init__(self, name: str):
        self.__name = name
        self.__handler = 0

    def is_open(self):
        if self.__handler:
            try:
                win32pipe.PeekNamedPipe(self.__handler, 0)
            except Exception:
                self.close()
                return False
            return True
        return False

    def open(self):
        if self.__handler:
            return False

        try:
            # Open pipe file
            self.__handler = win32file.CreateFile('\\\\.\\pipe\\' + self.__name,
                                                  win32file.GENERIC_READ | win32file.GENERIC_WRITE, 0, None,
                                                  win32file.OPEN_EXISTING, win32file.FILE_ATTRIBUTE_NORMAL, None)
            if not self.__handler:
                err = win32api.GetLastError()
                raise Exception('Failed to open pipe [error code: %s]' % err)

            # Set pipe mode
            win32pipe.SetNamedPipeHandleState(self.__handler, win32pipe.PIPE_READMODE_MESSAGE, None, None)
        except Exception as ex:
            self.close()
            raise ex
        return True

    def close(self):
        if self.__handler:
            try:
                win32api.CloseHandle(self.__handler)
                self.__handler = 0
            except Exception:
                pass

    def has_data(self):
        if self.__handler:
            try:
                (data, nAvail, nMessage) = win32pipe.PeekNamedPipe(self.__handler, 0)
                return nAvail > 0
            except Exception:
                self.close()
                pass
        return False

    def read(self):
        if self.__handler:
            try:
                (read, nAvail, nMessage) = win32pipe.PeekNamedPipe(self.__handler, 0)
                if nAvail > 0:
                    (err, data) = win32file.ReadFile(self.__handler, nAvail, None)
                    if err == 0:
                        return data
            except Exception:
                self.close()
                return False
        return None


class PipelineWriter:
    def __init__(self, name: str, buffer_length=4096):
        self.__name = name
        self.__buffer_length = buffer_length
        self.__handler = 0
        self.__wait_for_connection = False
        self.__thread_accept_connection = None
        self.__connected = False

    def is_open(self):
        if self.__handler:
            if self.__connected:
                try:
                    win32pipe.PeekNamedPipe(self.__handler, 0)
                except Exception:
                    self.close()
                    return False
            return True
        return False

    def open(self):
        if self.__handler:
            return False

        try:
            # Create pipe line
            self.__handler = win32pipe.CreateNamedPipe('\\\\.\\pipe\\' + self.__name,
                                                       win32pipe.PIPE_ACCESS_DUPLEX | win32file.FILE_FLAG_OVERLAPPED,
                                                       win32pipe.PIPE_TYPE_MESSAGE |
                                                       win32pipe.PIPE_READMODE_MESSAGE |
                                                       win32pipe.PIPE_WAIT,
                                                       1, self.__buffer_length, self.__buffer_length, 0, None)
            if not self.__handler:
                err = win32api.GetLastError()
                raise Exception('Failed to create pipe [error code: %s]' % err)

            # Wait for connection in async mode
            self.__wait_for_connection = True
            self.__thread_accept_connection = threading.Thread(target=self.__accept_connection)
            self.__thread_accept_connection.start()
        except Exception as ex:
            self.close()
            raise ex

        return True

    def close(self):
        try:
            self.__wait_for_connection = False
            if self.__thread_accept_connection:
                if self.__thread_accept_connection.is_alive():
                    self.__thread_accept_connection.join()
                self.__thread_accept_connection = None
            if self.__handler:
                if self.__connected:
                    win32pipe.DisconnectNamedPipe(self.__handler)
                win32api.CloseHandle(self.__handler)
                self.__handler = 0
            self.__connected = False
        except Exception:
            pass

    def wait_for_reader(self, timeout=None):
        if self.__wait_for_connection and self.__thread_accept_connection:
            if timeout is None:
                if self.__thread_accept_connection.is_alive():
                    self.__thread_accept_connection.join()
            else:
                while timeout > 0 and not self.__connected:
                    time.sleep(1)
                    timeout = timeout - 1

    def has_reader(self):
        return self.is_open() and self.__connected

    def write(self, data: bytes):
        if self.__handler and self.__connected:
            try:
                (err, bytes_written) = win32file.WriteFile(self.__handler, data)
                if bytes_written > 0:
                    return True
            except Exception:
                self.close()
                pass
        return False

    def __accept_connection(self):
        try:
            while self.__wait_for_connection and self.__handler:
                ret = win32pipe.ConnectNamedPipe(self.__handler, pywintypes.OVERLAPPED())
                if ret == winerror.ERROR_PIPE_CONNECTED:
                    self.__connected = True
                    break
                elif ret != winerror.ERROR_IO_PENDING:
                    win32api.CloseHandle(self.__handler)
                    self.__handler = 0
                    break
                time.sleep(1)
        except Exception:
            if self.__handler:
                win32api.CloseHandle(self.__handler)
                self.__handler = 0
        self.__wait_for_connection = False
