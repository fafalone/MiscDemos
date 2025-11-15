# MiscDemos
Old VB6 demos of mine converted to tB with no additional features/notes.\
All projects have been updated to use WinDevLib instead of oleexp.tlb and/or local API defs, and to support x64 builds.

**IFileDialog64** - This shows conversion of my VB6 project [[VB6] Using the new IFileDialog interface for customizable Open/Save](https://www.vbforums.com/showthread.php?786031-VB6-Using-the-new-IFileDialog-interface-for-customizable-Open-Save-(TLB-Vista-)) at each stage in converting to x64: The original VB6 import that works as-is with no modification, switching definitions to WinDevLib, updating for x64 support, and finally eliminating unncessary v-table swapping by using new tB features.

**ITaskbarListDemo** - Show an overlay icon, progress bar, custom thumbnail with buttons, and more on your app's Taskbar item. A port of [[VB6] Win7 Taskbar Features with ITaskbarList3 : overlay, progress in taskbar, etc](https://www.vbforums.com/showthread.php?786173).

**ShellLibrarySample** - How to use the features of Win7+ Library folders. A port of [[VB6] Working with Libraries (Win7+)](https://www.vbforums.com/showthread.php?785423)

**SendToDemo** - Shows how to pop up the Send To context menu for files. A port of [[VB6, Vista+] Add the Windows Send To submenu to your popup menu](https://www.vbforums.com/showthread.php?839949-VB6-Vista-Add-the-Windows-Send-To-submenu-to-your-popup-menu)

**MP3CoverArt** - Add album art to mp3 files. A port of [[VB6] Write MP3 Album Art and other tags using the Windows Property System](https://www.vbforums.com/showthread.php?880337-VB6-Write-MP3-Album-Art-and-other-tags-using-the-Windows-Property-System)

**VirtualFileDragDrop** - Drag and drop files that aren't created until actually dropped. A port of [[VB6] Virtual File Drag Drop](https://www.vbforums.com/showthread.php?866503-VB6-Virtual-File-Drag-Drop)

**cNamespaceWalk** - A class wrapper for INamespaceWalk, lists folder contents by shell order with levels and cancel options. A port of [[VB6] List files by level from a folder, in natural sorted order using INamespaceWalk](https://www.vbforums.com/showthread.php?794171-VB6-List-files-by-level-from-a-folder-in-natural-sorted-order-using-INamespaceWalk). Some important lessons for x64 conversion: Switched comctl32.ocx for VBCCR, switching to WDL meant some changes in the class since the VB-hack of needing double implementations of the callback isn't preserved, and an important reminder about the pitfalls of taking shortcuts with String pointers... instead of just passing StrPtr(string), `CoTaskMemAlloc` was needed. 

**BrowseForFolderFilter** - Shows how to filter the contents of `SHBrowseForFolder` dialog using `IFolderFilter`. A port of [[VB6] SHBrowseForFolder - Custom filter for shown items: BFFM_IUNKNOWN/IFolderFilter](https://www.vbforums.com/showthread.php?839997-VB6-SHBrowseForFolder-Custom-filter-for-shown-items-BFFM_IUNKNOWN-IFolderFilter). Added minor new features to toggle whether .zip/.cab files count as folders and use of anchors to make all form elements resize/move without new code.
