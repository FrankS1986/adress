   M S B A M L   `   `   `  ����E  AAdressbuch, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null'clr-namespace:Adressbuch
Adressbuch  5   	       Adressbuch.MainWindow   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   *localclr-namespace:Adressbuch     5   	     � Title$  
Adressbuch��5	   	   $	��600��6   $
��1000��6,   .��5      �  � ColumnDefinitions 5   
   �� 5      $	��300 �6   �� 5       � RowDefinitions 5   
   �� 5      $��50 �6   �� 5      $��26 �6   �� 5      �� 5      $��70 �6   .+�5   
   �� ���	#FF6FA8DC$��0��6,   $��2��6:   $��Center=�6P   .+�5      �� �� 5      $��Adressbuch ��6   $��30(�6*   5       �� 5      $��lupe.png��6   $��22��6(   $��20��64   $��
340,0,0,15q�6?   6W    �� 5!      -   5         ��Name$ suche��5!      $��22��6/   $	��300��6;   $��3,0,0,15q�6G   $�� ��6Y   $��Italic$�6�   5(      5)      �� 5*   
   $��1��6   $��1��6"   $��2��62   $��0,0,0,5q�6C   .+�5+      �� -   5         ��Name$ Kontakte��5+      .��5-      &   System.Windows.Controls.Calendar  -   5          Name$ calendar��5-      $y�Visible=�61    J� RenderTransformOrigin$ 0.5,0.5&�6F   $	��182��6f   z�5.      V� 5/      .~�50      ��  �� ScaleX$	 2.5Q�6-    �� ScaleY$ 2Q�6:   �� 51      �� 52      T� 53      	 T� Y$	 90Q�61   
 T� X$
 -2Q�68   54      58      5@      5B      �� 5C   
   $��1��6   ���	#FFBBB2B2.+�5D      �� $��Right=�6   �� 5F      -   5         ��Name$ gebBild��5F      $��geb.png��6)   $��25��6:   $��25��6F   $
��Fill=�6r    � ForceCursor$
 True��6�   5H       �� 5I      -   5        $ kontakteBild��5I      $��kontakte.png��6.   $��25��6D   $��25��6P   $
��Fill=�6~   6�    �� 5J      -   5        $ 
deleteBild��5J      $��
delete.png��6,   $��25��6@   $��25��6L   $
��Fill=�6x   6�    �� 5K      -   5        $ addBild��5K      $��add.png��6)   $��25��6;   $��25��6G   $	��3,0q�6q   $
��Fill=�6~   6�   5P      5S      �� 5T   
   $��3��6   $��1��6"   .+�5U      �� $��63��6   $
��Left=�6$   �� 5W      -   5        $ bebBild��5W      $��/geb.png��6%   $��50��67   $��50��6C   5Y       �� 5[      .��6   Nächster Geburtstag:60    �� 5\      -	   5         ��Name$ gebnext��5\      6+    �� 5]      -
   5        $ 
bearbeiten��5]      $��/bearbeiten.png��6+   $��0��6D   $��0��6O   $��	200,0,0,0q�6�   5_      5a      5b      �� 5d   
   $��2��6   $��2��6"   .+�5f      �� -   5         ��Name$ lboxKontakte��5f      $	��370��6-   6:   5h      5i      5j      