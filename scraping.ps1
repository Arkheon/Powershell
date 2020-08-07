#https://www.warcrafttavern.com/loot/molten-core/?wpv_view_count=5976-TCPID5970&wpv_post_search=&wpv-loot-class=0&wpv-loot-class-role=0&wpv-loot-acquisition=0&wpv-wpcf-patch-added=&wpv-wpcf-item-quality=&wpv-wpcf-item-slot=&wpv-wpcf-item-type=&wpv-wpcf-level_min=&wpv-wpcf-level_max=&wpv-loot-resistance=0&wpv_paged=1

$tableau_adresse = 1

foreach ($adresse_ID in $tableau_adresse )
{
$adresseMC = "https://www.warcrafttavern.com/loot/molten-core/?wpv_view_count=5976-TCPID5970&wpv_post_search=&wpv-loot-class=0&wpv-loot-class-role=0&wpv-loot-acquisition=0&wpv-wpcf-patch-added=&wpv-wpcf-item-quality=&wpv-wpcf-item-slot=&wpv-wpcf-item-type=&wpv-wpcf-level_min=&wpv-wpcf-level_max=&wpv-loot-resistance=0&wpv_paged=$adresse_ID"
invoke-webrequest -Uri "$adresseMC" -outfile "html\mc_$adresse_ID.html"
 
$html = New-Object -ComObject "HTMLFile"
 
$html.IHTMLDocument2_write($(Get-Content .\html\mc_$adresse_ID.html -raw))
 
#$html.getElementsByTagName('div')| Where {$_.getAttributeNode('class').Value -eq 'mtr-cell-content'} | select InnerHtml

#$html.getElementsByTagName('div')| Where {$_.getAttributeNode('class').Value -eq 'mtr-cell-content'} | select innerText 

$a = $html.getElementsByTagName('div')| Where {$_.getAttributeNode('class').Value -eq 'mtr-cell-content'} | Select-Object -Expand innerText | Out-String | Add-Content test.csv


<# for ($j=1;$j -le 14;$j++){
$f=for ($i=0;$i -lt $a.count;$i+=1) { 
$a[$i] | Out-String | Add-Content test.csv
}
}
$f #>

#req nom item
#$html.getElementsByTagName('td') | Where {$_.getAttributeNode('class').Value -Like '*Item*'} | select innerText
#retourne : <DIV class=mtr-cell-content><A href="https://vanillawowdb.com/?item=18399" rel="noopener noreferrer" target=_blank>Ocean's Breeze</A></DIV>

#test : retourne ID
#$html.getElementsByTagName('td') |  Where-Object { $_.className -Like '*Item*' } | ForEach-Object { $_.getElementsByTagName('a') } |   Select-Object -Expand href | ForEach-Object {$_.Substring($_.Length - 5)}
#$html.getElementsByTagName('td') |  Where-Object { $_.className -Like '*Item*' } | select OuterText 



#zone
#$html.getElementsByTagName('td') | Where {$_.getAttributeNode('name').Value -Like '*Zone*'}
#$html.getElementsByTagName('td') |  Where-Object { $_.className -Like '*zone*' } | select innerText 
#loot acquisition
#$html.getElementsByTagName('td') | Where {$_.getAttributeNode('class').Value -Like '*Loot Acquisition*'}
#Acquisition Name
#$html.getElementsByTagName('td') | Where {$_.getAttributeNode('class').Value -Like '*Acquisition Name*'}
#item slot 
#$html.getElementsByTagName('td') | Where {$_.getAttributeNode('class').Value -Like '*Item Slot*'}
#item type
#$html.getElementsByTagName('td') | Where {$_.getAttributeNode('class').Value -Like '*Item Type*'}
}


<# <tr>
<td class="loot-list-item mtr-td-tag" data-mtr-content="Item"><div class="mtr-cell-content"><a href="https://vanillawowdb.com/?item=18872" target="_blank" rel="noopener noreferrer">Manastorm Leggings</a></div></td>
<td class="loot-table-center mtr-td-tag" data-mtr-content="Level"><div class="mtr-cell-content">60</div></td>
<td data-mtr-content="Zone" class="mtr-td-tag"><div class="mtr-cell-content">Molten Core</div></td>
<td data-mtr-content="Loot Acquisition" class="mtr-td-tag"><div class="mtr-cell-content">Boss</div></td>
<td data-mtr-content="Acquisition Name" class="mtr-td-tag"><div class="mtr-cell-content"><a href="" target="_blank" rel="noopener noreferrer">All Bosses</a></div></td>
<td data-mtr-content="Item Slot" class="mtr-td-tag"><div class="mtr-cell-content">Legs</div></td>
<td data-mtr-content="Item Type" class="mtr-td-tag"><div class="mtr-cell-content">Cloth</div></td>
<td class="loot-table-center mtr-td-tag" data-mtr-content="Drop Chance"><div class="mtr-cell-content">2.17%</div></td>
</tr> #>
