<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:xs="http://www.w3.org/2001/XMLSchema">
<!--<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="2.0">-->
    <xsl:output method="html" encoding="utf-8"/>

    <xsl:template match="/">
        <html>
            <meta charset="utf-8"></meta>
        <body>
        <table>
            <thead>
            <tr>
                <th width="150">code</th>
                <th width="150">parent code</th>
                <th>name</th>
                <th>edition</th>
                <th>header</th>
                <th>note</th>
                <th>notePosition</th>
                <th>indexBefore</th>
                <th>indexAfter</th>
                <th>class</th>
            </tr>
            </thead>
            <tbody>
        <xsl:for-each select="//Section">
            <tr>
                <td><xsl:value-of select="IpcEntry/@Symbol"/></td>
                <td/>
                <td><xsl:value-of select="IpcEntry/Title/text()"/></td>
                <td><xsl:value-of select="IpcEntry/@Edition"/></td>
                <td><xsl:value-of select="Header/Title/text()"/></td>
                <td/>
                <td/>
                <td/>
                <td/>
                <td><xsl:value-of select="local-name()"/></td>
            </tr>

            <xsl:for-each select="Class">
                <tr>
                    <td><xsl:value-of select="IpcEntry/@Symbol"/></td>
                    <td><xsl:value-of select="parent::node()/IpcEntry/@Symbol"/></td>
                    <td><xsl:value-of select="IpcEntry/Title/text()"/></td>
                    <td><xsl:value-of select="IpcEntry/@Edition"/></td>
                    <td><xsl:value-of select="Header/Title/text()"/></td>
                    <td/>
                    <td/>
                    <td/>
                    <td/>
                    <td><xsl:value-of select="local-name()"/></td>
                </tr>

                <xsl:for-each select="SubClass">
                    <tr>
                        <td><xsl:value-of select="IpcEntry/@Symbol"/></td>
                        <td><xsl:value-of select="parent::node()/IpcEntry/@Symbol"/></td>
                        <td><xsl:value-of select="IpcEntry/Title/text()"/></td>
                        <td><xsl:value-of select="IpcEntry/@Edition"/></td>
                        <td><xsl:value-of select="Header/Title/text()"/></td>
                        <td/>
                        <td/>
                        <td/>
                        <td/>
                        <td><xsl:value-of select="local-name()"/></td>
                    </tr>

                    <xsl:for-each select="MainGroup">
                        <tr>
                            <td>
<!--                                <xsl:analyze-string select="IpcEntry/@Symbol" regex="\w+">-->
<!--                                    <xsl:matching-substring>-->
<!--                                        <xsl:value-of select="."/>-->
<!--                                    </xsl:matching-substring>-->
<!--                                    <xsl:non-matching-substring>-->
<!--                                        test-->
<!--                                    </xsl:non-matching-substring>-->
<!--                                </xsl:analyze-string>-->
                                <xsl:value-of select="IpcEntry/@Symbol"/>
                            </td>
                            <td><xsl:value-of select="parent::node()/IpcEntry/@Symbol"/></td>
                            <td><xsl:value-of select="IpcEntry/Title/text()"/></td>
                            <td><xsl:value-of select="IpcEntry/@Edition"/></td>
                            <td><xsl:value-of select="Header/Title/text()"/></td>
                            <td/>
                            <td/>
                            <td/>
                            <td/>
                            <td><xsl:value-of select="local-name()"/></td>
                            <!--                        <td><xsl:value-of select="preceding-sibling::*[1][self::head]"/></td>-->
                        </tr>

                        <xsl:for-each select="SubGroup">
                            <tr>
                                <td><xsl:value-of select="IpcEntry/@Symbol"/></td>
                                <td><xsl:value-of select="parent::node()/IpcEntry/@Symbol"/></td>
                                <td><xsl:value-of select="IpcEntry/Title/text()"/></td>
                                <td><xsl:value-of select="IpcEntry/@Edition"/></td>
                                <td><xsl:value-of select="Header/Title/text()"/></td>
                                <td/>
                                <td/>
                                <td/>
                                <td/>
                                <td><xsl:value-of select="local-name()"/></td>
                            </tr>
                        </xsl:for-each>
                    </xsl:for-each>
                </xsl:for-each>

            </xsl:for-each>
        </xsl:for-each>
            </tbody>
        </table>
        </body>
        </html>
    </xsl:template>

<!--</xsl:transform>-->
</xsl:stylesheet>