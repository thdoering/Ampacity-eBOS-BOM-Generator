from typing import Any, Callable, List, Optional
from dataclasses import dataclass
from copy import deepcopy

@dataclass
class UndoState:
    """Represents a single state in the undo/redo stack"""
    state: Any
    description: str

class UndoManager:
    def __init__(self, max_states: int = 50):
        self.undo_stack: List[UndoState] = []
        self.redo_stack: List[UndoState] = []
        self.max_states = max_states
        self._state_getter: Optional[Callable[[], Any]] = None
        self._state_setter: Optional[Callable[[Any], None]] = None
        
    def set_callbacks(self, 
                     get_state: Callable[[], Any],
                     set_state: Callable[[Any], None]):
        """Set the callbacks for getting and setting state"""
        self._state_getter = get_state
        self._state_setter = set_state
        
    def push_state(self, description: str):
        """Push current state onto undo stack"""
        if not self._state_getter:
            raise RuntimeError("State getter callback not set")
            
        current_state = deepcopy(self._state_getter())
        self.undo_stack.append(UndoState(current_state, description))
        self.redo_stack.clear()  # Clear redo stack when new state is pushed
        
        # Maintain maximum stack size
        if len(self.undo_stack) > self.max_states:
            self.undo_stack.pop(0)
            
    def undo(self) -> Optional[str]:
        """Undo last action, returns description of undone action"""
        if not self.undo_stack or not self._state_setter:
            return None
                
        # Get the state we want to return to
        last_state = self.undo_stack.pop()
        
        # Get current state before applying the change
        current_state = deepcopy(self._state_getter())
        
        # Push current state to redo stack
        self.redo_stack.append(UndoState(current_state, last_state.description))
        
        # Restore previous state
        self._state_setter(last_state.state)
        
        return last_state.description
        
    def redo(self) -> Optional[str]:
        """Redo last undone action, returns description of redone action"""
        if not self.redo_stack or not self._state_setter:
            return None
            
        # Save current state to undo stack
        current_state = deepcopy(self._state_getter())
        next_state = self.redo_stack.pop()
        self.undo_stack.append(UndoState(current_state, next_state.description))
        
        # Restore next state
        self._state_setter(next_state.state)
        return next_state.description
        
    def can_undo(self) -> bool:
        """Check if undo is available"""
        return len(self.undo_stack) > 0
        
    def can_redo(self) -> bool:
        """Check if redo is available"""
        return len(self.redo_stack) > 0