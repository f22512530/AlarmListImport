   M S B A M L   `   `   `  ����F  BExcelImport, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null)clr-namespace:ExcelImportExcelImport  5   	       ExcelImport.MainWindow   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   +localclr-namespace:ExcelImport     5   	     � Title$  
MainWindow��5   	   $	��450��6   $	��800��6)   .��5	      �  � RowDefinitions 5
   
   �� 5      $
��0.2* �6   �� 5      $
��0.2* �6   �� 5      .+�5   
   �� $��0��6    �� Orientation$ 
Horizontal=�6"   $
��20,0q�6;   .+�5      �� -   5        $��CreateDTBtn��5      $��建立資料庫��6(   $	��100��68   $��30��6D   $
��Left=�6P   $��Center=�6k   6�   �� 5      -   5        $��CreateDTRemind��5      $
��Left=�6.   $��Center=�6I   $
��20,0q�6d   5      �� 5   
   $��1��6   $ 
Horizontal=�6"   $
��20,0q�6;   .+�5      �� -   5        $��fileLoad��5      $	��100��6%   $��30��61   $��載入��6=   $
��Left=�5      $��Center=�6,   �� 5      -   5        $��dataPath��5      $
��Auto��6&   $��30��63   $	��200��6?   $	��ABC��6N   {�.  $��Center=�5      $��Center=�6/   $��20,0,0,0q�6J   5      �� 5   
   -   5        $��InsertOverflow��5      $��2��6(   $��20q�65    �� 
IsReadOnly . 6A   5      5      