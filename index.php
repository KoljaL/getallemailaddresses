<html>

<link rel="stylesheet" href="style.css"/>
<script type="text/javascript">function selectElementContents(e){var t,n,o=document.body;if(document.createRange&&window.getSelection){t=document.createRange(),(n=window.getSelection()).removeAllRanges();try{t.selectNodeContents(e),n.addRange(t)}catch(o){t.selectNode(e),n.addRange(t)}document.execCommand("copy")}else o.createTextRange&&((t=o.createTextRange()).moveToElementText(e),t.select(),t.execCommand("Copy"))}</script>

<head>
<title>Extract email addresses from inbox using IMAP</title>
<meta description="Script for extracting email addresses from messages' headers and bodies from any inbox using IMAP into CSV (Excel)">
</head>

<body>
<h1>Extract email addresses from messages' headers and bodies from any inbox using IMAP</h1>
<a href="https://stackoverflow.com/questions/48217519/extract-email-addresses-list-from-inbox-using-php-and-imap/48219775#48219775" style="color: grey; font-size: 16px; text-decoration: none;">based on great code by Eugene Lycenok</a>
<br><br><br><br>


<form method="GET" class="card card-block bg-faded" autocomplete="off">

  <div class="form-group input-group">
    <label class="has-float-label">
      <input class="form-control" name="mailbox" type="text" size="50"    placeholder="{imap.gmail.com:993/imap/ssl}INBOX"/ >
      <span>IMAP Mailbox</span>
    </label>

    <label class="has-float-label" >
      <input class="form-control"  name="login" type="text"  size="10"   placeholder="Arthur_Dent"/ size="30">
      <span>Login Name</span>
    </label>

    <label class="has-float-label" >
      <input class="form-control" name="password" type="password" size="10"   autocomplete="off" placeholder="&#9679;&#9679;&#9679;&#9679;&#9679;&#9679;&#9679;&#9679;&#9679;&#9679;"/ >
      <span>Passwort</span>
    </label>

    <label class="has-float-label" >
      <input class="form-control" name="maxcount" type="text" size="7" placeholder="42"/ value=<?php echo ((isset($_REQUEST['maxcount'])) ?  $_REQUEST['maxcount'] : '42'); ?>>
      <span>Anzahl</span>
    </label>
  </div>
<br>

  <div class="ck-button"><label><input type="checkbox" name="nummer" value="1" <?php echo ((isset($_REQUEST['nummer'])) ? ' checked="checked" ' : ''); ?> ><span>Nummer</span></label></div>
  <div class="ck-button"><label><input type="checkbox" name="datum" value="1" <?php echo ((isset($_REQUEST['datum'])) ? ' checked="checked" ' : ''); ?> ><span>Datum</span></label></div>
  <div class="ck-button"><label><input type="checkbox" name="betreff" value="1" <?php echo ((isset($_REQUEST['betreff'])) ? ' checked="checked" ' : ''); ?> ><span>Betreff</span></label></div>
  <div class="ck-button"><label><input type="checkbox" name="vonname" value="1" <?php echo ((isset($_REQUEST['vonname'])) ? ' checked="checked" ' : ''); ?> ><span>von Name</span></label></div>
  <div class="ck-button"><label><input type="checkbox" name="vonemail" value="1" <?php echo ((isset($_REQUEST['vonemail'])) ? ' checked="checked" ' : ''); ?> ><span>von Email</span></label></div>
  <div class="ck-button"><label><input type="checkbox" name="anname" value="1" <?php echo ((isset($_REQUEST['anname'])) ? ' checked="checked" ' : ''); ?> ><span>an Name</span></label></div>
  <div class="ck-button"><label><input type="checkbox" name="anemail" value="1" <?php echo ((isset($_REQUEST['anemail'])) ? ' checked="checked" ' : ''); ?> ><span>an Email</span></label></div>
  <div class="ck-button"><label><input type="checkbox" name="alleadressen" value="1" <?php echo ((isset($_REQUEST['alleadressen'])) ? ' checked="checked" ' : ''); ?> ><span>Alle Adressen</span></label></div>

  <input name="submitFlag" type="hidden" value="1"></input>
  <button class="btn panic" type="submit">Don't Panik</button>

</form>

<?php

if (isset($_REQUEST['submitFlag']))
    {

        define("MAX_EMAIL_COUNT", $_REQUEST['maxcount']);

        /* took from https://gist.github.com/agarzon/3123118 */
        function extractEmail($content)
            {
                $regexp = '/([a-z0-9_\.\-])+\@(([a-z0-9\-])+\.)+([a-z0-9]{2,4})+/i';
                preg_match_all($regexp, $content, $m);
                return isset($m[0]) ? $m[0] : array();
            }

        function getAddressText(&$emailList, &$nameList, $addressObject)
            {

                $emailList = '';
                $nameList  = '';

                foreach ($addressObject as $object)
                    {
                        $emailList .= ';';

                        if (isset($object->personal))
                            {
                                $emailList .= $object->personal;
                            }

                        $nameList .= ';';
                        if (isset($object->mailbox) && isset($object->host))
                            {
                                $nameList .= $object->mailbox . "@" . $object->host;
                            }
                    }

                $emailList     = ltrim($emailList, ';');
                $nameList      = ltrim($nameList, ';');
            }

        //$allAdresses = '';
        function processMessage($mbox, $messageNumber)
            {

                // get imap_fetch header and put single lines into array
                $header        = imap_rfc822_parse_headers(imap_fetchheader($mbox, $messageNumber));

                $fromEmailList = '';
                $fromNameList  = '';
                if (isset($header->from))
                    {
                        getAddressText($fromEmailList, $fromNameList, $header->from);
                    }

                $subject     = iconv_mime_decode($header->subject, 0, "ISO-8859-1"); // $header->subject;
                $udate       = '';
                $udate       = strtotime($header->date);
                $udate       = date("d.m.Y H:i", $udate);

                $toEmailList = '';
                $toNameList  = '';
                if (isset($header->to))
                    {
                        getAddressText($toEmailList, $toNameList, $header->to);
                    }

                $fromEmailList = str_replace('"', '', iconv_mime_decode($fromEmailList, 0, "ISO-8859-1"));
                $toEmailList   = str_replace('"', '', iconv_mime_decode($toEmailList, 0, "ISO-8859-1"));

                if (isset($_REQUEST['alleadressen']))
                    {
                        $body          = imap_fetchbody($mbox, $messageNumber, 1);
                        $bodyEmailList = implode(';', extractEmail($body));
                    }

                echo "<tr>";
                if (isset($_REQUEST['nummer']))
                    {
                        echo "<td>" . $messageNumber . "</td>";
                    }
                if (isset($_REQUEST['datum']))
                    {
                        echo "<td>" . $udate . "</td>";
                    }
                if (isset($_REQUEST['betreff']))
                    {
                        echo "<td>" . $subject . "</td>";
                    }
                if (isset($_REQUEST['vonname']))
                    {
                        echo "<td>" . $fromEmailList . "</td>";
                    }
                if (isset($_REQUEST['vonemail']))
                    {
                        echo "<td>" . $fromNameList . "</td>";
                    }
                if (isset($_REQUEST['anname']))
                    {
                        echo "<td>" . $toEmailList . "</td>";
                    }
                if (isset($_REQUEST['anemail']))
                    {
                        echo "<td>" . $toNameList . "</td>";
                    }
                if (isset($_REQUEST['alleadressen']))
                    {
                        echo "<td>" . $bodyEmailList . "</td>";
                    }
                echo "</tr> \n";

                // $allAdresses .= $bodyEmailList;

            }

        $beginn   = microtime(true);

        // imap_timeout(IMAP_OPENTIMEOUT, 300);
        // Open pop mailbox
        if (!$mbox     = imap_open($_REQUEST['mailbox'], $_REQUEST['login'], $_REQUEST['password']))
            {
                die('Cannot connect/check pop mail! Exiting');
            }
        if ($hdr      = imap_check($mbox))
            {
                $msgCount = $hdr->Nmsgs;
            }
        else
            {
                echo "Failed to get mail";
                exit;
            }


        echo "<table class='table' id='tableId'> \n";
        echo "  <tr>";
        if (isset($_REQUEST['nummer']))
            {
                echo  "<td style='font-weight:bold'>Nr</td>";
            }
        if (isset($_REQUEST['datum']))
            {
                echo "<td style='font-weight:bold'>Datum</td>";
            }
        if (isset($_REQUEST['betreff']))
            {
                echo "<td style='font-weight:bold'>Betreff</td>";
            }
        if (isset($_REQUEST['vonname']))
            {
                echo "<td style='font-weight:bold'>von Name</td>";
            }
        if (isset($_REQUEST['vonemail']))
            {
                echo "<td style='font-weight:bold'>von Email</td>";
            }
        if (isset($_REQUEST['anname']))
            {
                echo "<td style='font-weight:bold'>an Name</td>";
            }
        if (isset($_REQUEST['anemail']))
            {
                echo "<td style='font-weight:bold'>an Email</td>";
            }
        if (isset($_REQUEST['alleadressen']))
            {
                echo "<td style='font-weight:bold'>extracted from body</td>";
            }
        echo "</tr> \n";


        for ($X = 1; $X <= min($msgCount, MAX_EMAIL_COUNT); $X++)
            {
                processMessage($mbox, $X);
            }

        imap_close($mbox);
        $dauer = microtime(true) - $beginn;



        echo "<div class='info_out'> \n";
        echo "  <div class='info'> Anzahl der Mails im Postfach = " . $msgCount .  "</div> \n";
        echo "  <div class='info'> Skriptlaufzeit: " . round($dauer, 2) . " Sekunden </div> \n\n";

        echo "  <div id='hide_if_no_JS' class='hide_if_no_JS_hide info'> \n";
        echo "    <input class=\"btn\" type=\"button\" value=\"Tabelle in die Zwischenablage kopieren\" onclick=\"selectElementContents( document.getElementById('tableId') );\"> \n";
        echo "  </div> \n";
        echo "  <script type=\"text/javascript\">document.getElementById(\"hide_if_no_JS\").className = \"hide_if_no_JS_show info \";</script> \n\n";

        echo "</div>  \n";

        echo "</table> \n\n";
    }
?>
</body>
</html>
