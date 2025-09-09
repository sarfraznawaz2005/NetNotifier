# NetNotifier Async Analysis and Implementation Plan

## Identified Synchronous Blocking Functions

### Network Operations (Primary Blocking Sources)
1. **PingSync** (line 520)
   - Uses `RunWait` for ping command execution
   - Blocks GUI during ping operations with retries
   - Called from: DetermineConnectionStatus, GetLastPingTime, TestConnectionFunc

2. **CheckDNS** (line 222)
   - Uses `RunWait` for nslookup command
   - Fallback to synchronous DllCall for gethostbyname
   - Called from: DetermineConnectionStatus

3. **CheckCaptivePortal** (line 611)
   - Uses synchronous WinHttpRequest.Send()
   - Blocks during HTTP request to google.com
   - Called from: DetermineConnectionStatus

4. **GetPublicIPFetch** (line 674)
   - Uses synchronous WinHttpRequest for IP fetching
   - Blocks during HTTP requests to api.ipify.org/icanhazip.com
   - Called from: HandleStatusChange, GetPublicIPCached

5. **GetDownloadSpeed** (line 654)
   - Uses synchronous WinHttpRequest for speed test
   - Blocks during file download
   - Not currently called in main flow

6. **CheckTCPConnection** (line 570)
   - Uses synchronous socket operations via DllCall
   - Blocks during TCP connection attempts
   - Not currently called in main flow

### System Operations
7. **GetDefaultGateway** (line 261)
   - Uses synchronous WMI queries
   - Called from: DetermineConnectionStatus

8. **CheckNetworkInterfaceStatus** (line 288)
   - Uses synchronous WMI queries
   - Called from: DetermineConnectionStatus

### File Operations
9. **LoadSettings** (line 66)
   - Uses synchronous IniRead operations
   - Called from: __New()

10. **SettingsButtonSave** (line 395)
    - Uses synchronous IniWrite operations
    - Called from: Settings GUI save button

## Root Cause Analysis

The main blocking function is **CheckConnection** (line 197) which:
- Prevents overlapping checks with `IsChecking` flag
- Calls **DetermineConnectionStatus** (line 149) synchronously
- All network checks happen sequentially in the main thread
- GUI freezes during the entire check cycle (potentially 10-30 seconds)

## Async Implementation Strategies

### Strategy 1: Background Process Architecture
- Create separate AHK script for network checks
- Main script communicates via files/messages
- Background script runs checks asynchronously
- Main script updates GUI based on background results

### Strategy 2: Timer-Based Async Operations
- Break network checks into smaller timed operations
- Use multiple timers for different check types
- Update GUI incrementally as results come in
- Implement timeout mechanisms for each operation

### Strategy 3: Asynchronous HTTP Requests
- Replace WinHttpRequest synchronous calls with async versions
- Use WinHttpRequest events for completion callbacks
- Implement proper error handling for async operations

### Strategy 4: Command-Line Async Execution
- Replace `RunWait` with `Run` and callback mechanisms
- Use temporary files for result communication
- Implement polling or file watching for completion

## Implementation Plan (Todos)

### Phase 1: Core Async Infrastructure
- [ ] Create AsyncPing class to replace PingSync
  - Use `Run` with callback script for ping execution
  - Implement result file-based communication
  - Add timeout handling for stuck processes

- [ ] Create AsyncHttp class for HTTP operations
  - Replace synchronous WinHttpRequest with async version
  - Implement callback system for request completion
  - Handle multiple concurrent requests

- [ ] Create AsyncWmi class for system queries
  - Move WMI operations to background threads
  - Implement caching to reduce frequency of calls
  - Add timeout protection

### Phase 2: Refactor Main Check Logic
- [ ] Refactor DetermineConnectionStatus to async
  - Break into smaller async steps
  - Use promise-like pattern with callbacks
  - Implement early exit on first successful check

- [ ] Update CheckConnection to use async pattern
  - Remove blocking `IsChecking` flag
  - Allow multiple concurrent checks if needed
  - Implement proper cleanup on app exit

- [ ] Create AsyncResultManager
  - Handle aggregation of multiple async results
  - Implement timeout for overall check operation
  - Provide fallback mechanisms

### Phase 3: GUI Responsiveness Improvements
- [ ] Update HandleStatusChange for incremental updates
  - Allow partial status updates during checks
  - Show "Checking..." status in tray tooltip
  - Implement progress indicators if needed

- [ ] Make Settings GUI non-blocking
  - Move validation to background if needed
  - Show loading indicators during saves
  - Prevent multiple simultaneous operations

- [ ] Implement AsyncTooltip updates
  - Update tooltip without blocking main thread
  - Cache expensive calculations
  - Use timers for periodic refreshes

### Phase 4: Error Handling and Reliability
- [ ] Add comprehensive timeout handling
  - Per-operation timeouts
  - Overall check timeout
  - Graceful degradation on failures

- [ ] Implement retry mechanisms
  - Exponential backoff for failed operations
  - Maximum retry limits
  - User-configurable retry settings

- [ ] Add async logging
  - Non-blocking log writes
  - Buffered logging to prevent I/O blocking
  - Log rotation without blocking

### Phase 5: Testing and Optimization
- [ ] Create async unit tests
  - Test each async operation independently
  - Verify GUI responsiveness during checks
  - Test timeout and error scenarios

- [ ] Performance optimization
  - Profile async operations
  - Optimize caching strategies
  - Minimize unnecessary operations

- [ ] Memory management
  - Clean up async resources properly
  - Prevent memory leaks in long-running operations
  - Monitor resource usage

## Migration Strategy

1. **Incremental Migration**: Start with PingSync replacement, then move to HTTP operations
2. **Backward Compatibility**: Maintain sync fallbacks during transition
3. **Feature Flags**: Allow users to enable/disable async features
4. **Rollback Plan**: Ability to revert to synchronous mode if issues arise

## Expected Benefits

- **GUI Responsiveness**: No more freezing during network checks
- **Better User Experience**: Real-time status updates
- **Improved Reliability**: Better timeout handling and error recovery
- **Scalability**: Ability to add more checks without performance impact
- **Maintainability**: Cleaner separation of concerns with async patterns

## Potential Challenges

- **AHK Limitations**: Limited built-in async support
- **Complexity**: Increased code complexity with async patterns
- **Debugging**: Harder to debug async operations
- **Resource Usage**: Potential increase in CPU/memory usage
- **Compatibility**: May require AHK v2 specific features

## Success Metrics

- GUI response time < 100ms during checks
- No blocking operations > 5 seconds
- Successful completion of all async operations
- Maintain current functionality and accuracy
- User-reported improvement in responsiveness