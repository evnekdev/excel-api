```mermaid
stateDiagram-v2
[*] --> Uninitialized
Uninitialized --> Initializing: xlAutoOpen/ensure
Initializing --> Initialized
Initialized --> Initialized: duplicate add/open
Initialized --> Closing: xlAutoClose
Closing --> Uninitialized
```
