from abc import ABC, abstractmethod
from typing import List

from excelutils.opserverpattern.Observer import Observer

"""
The Subject interface declares a set of methods for managing subscribers.
"""


class Subject(ABC):
    """
   The Subject owns some important state and notifies observers when the state
   changes.
   """

    _state: int = None
    """
    For the sake of simplicity, the Subject's state, essential to all
    subscribers, is stored in this variable.
    """

    _observers: List[Observer] = []
    """
    List of subscribers. In real life, the list of subscribers can be stored
    more comprehensively (categorized by event type, etc.).
    """

    def attach(self, observer: Observer) -> None:
        self._observers.append(observer)

    def detach(self, observer: Observer) -> None:
        self._observers.remove(observer)

    def notify(self):
        """
        Notify all observers about an event.
        """
        for observer in self._observers:
            observer.update()
