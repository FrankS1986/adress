   M S B A M L   `   `   `  ����E  AAdressbuch, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null'clr-namespace:Adressbuch
Adressbuch  5   	       Adressbuch.MainWindow   5   	   W SPresentationCore, Version=6.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=6.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=6.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=6.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=6.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   *localclr-namespace:Adressbuch     5   	     � Title$  
MainWindow��5	   	   $	��600��6   $
��1000��6,   .��5      �  � ColumnDefinitions 5   
   �� 5      $	��300 �6   �� 5       � RowDefinitions 5   
   �� 5      $��50 �6   �� 5      $��26 �6   �� 5      �� 5      $��70 �6   .+�5   
   �� ���	#FF6FA8DC$��0��6,   $��2��6:   $��Center=�6P   .+�5      �� �� 5      $��Adressbuch ��6   $��30(�6*   5       �� 5      $��lupe.png��6   $��22��6(   $��20��64   $��
340,0,0,15q�6?   6U    �� 5       -   5         ��Name$ suche��5       $��22��6'   $	��300��63   $��3,0,0,15q�6?   $�� ��6Q   .��6y   5#      5&      5'      �� 5(   
   $��1��6   $��1��6"   $��2��62   $��0,0,0,5q�6C   .+�5)      �� -   5         ��Name$ Kontakte��5)      .��5+      &   System.Windows.Controls.Calendar  -   5          Name$ calendar��5+      $y�Visible=�61    J� RenderTransformOrigin$ 0.5,0.5&�6F   $	��182��6f   z�5,      V� 5-      .~�5.      ��  �� ScaleX$	 2.5Q�6-    �� ScaleY$ 2Q�6:   �� 5/      �� 50      T� 51      	 T� Y$	 90Q�61   
 T� X$
 -2Q�68   52      56      5>      5@      �� 5A   
   $��1��6   ���	#FFBBB2B2.+�5B      �� $��Right=�6   �� 5D      -   5         ��Name$ gebBild��5D      $��geb.png��6)   $��25��6:   $��25��6F   $
��Fill=�6r    � ForceCursor$
 True��6�   5F       �� 5G      -   5        $ kontakteBild��5G      $��kontakte.png��6.   $��25��6D   $��25��6P   $
��Fill=�6~   6�    �� 5H      -   5        $ 
deleteBild��5H      $��
delete.png��6,   $��25��6@   $��25��6L   $
��Fill=�6x   6�    �� 5I      -   5        $ addBild��5I      $��add.png��6)   $��25��6;   $��25��6G   $	��3,0q�6q   $
��Fill=�6~   6�   5N      5Q      �� 5R   
   $��3��6   $��1��6"   .+�5S      �� $��63��6   $
��Left=�6$   �� 5U      -   5        $ bebBild��5U      $��/geb.png��6%   $��50��67   $��50��6C   5W       �� 5Y      .��6   Nächster Geburtstag:60    �� 5Z      -	   5         ��Name$ gebnext��5Z      6+    �� 5[      -
   5        $ 
bearbeiten��5[      $��/bearbeiten.png��6+   $��0��6D   $��0��6O   $��	200,0,0,0q�6�   5]      5_      5`      �� 5b   
   $��2��6   $��2��6"   .+�5d      �� -   5         ��Name$ lboxKontakte��5d      $	��370��6-   6:   5f      5g      5h      