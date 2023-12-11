from typing import Optional, List
import win32com.client as win32


class PowerPoint:
    def __init__(self, path: str) -> None:
        self.app = win32.gencache.EnsureDispatch('PowerPoint.Application')
        self.presentation = self.app.Presentations.Open(path)

    
    def save(self) -> None:
        self.presentation.Save()

    
    def close(self) -> None:
        self.save()
        self.presentation.Close()

    
    def exit(self) -> None:
        
        self.app.Quit()
		
    
    def slides(self) -> List[object]:
	    return self.presentation.Slides
	

    def slide(self, index) -> object:
	    return self.presentation.Slides(index)