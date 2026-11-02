<?php

/**
 * This file is part of PHPWord - A pure PHP library for reading and writing
 * word processing documents.
 *
 * PHPWord is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPWord/contributors.
 *
 * @see         https://github.com/PHPOffice/PHPWord
 *
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpWordTests\Reader\Word2007;

use PhpOffice\PhpWord\Element\Comment;
use PhpOffice\PhpWord\Element\Text;
use PhpOffice\PhpWord\Element\TextRun;
use PhpOffice\PhpWordTests\AbstractTestReader;

/**
 * Test class for PhpOffice\PhpWord\Reader\Word2007\Comments.
 */
class CommentsTest extends AbstractTestReader
{
    /**
     * A comment is linked to the elements delimited by its range, and not to the
     * element holding its reference.
     */
    public function testCommentRangeIsLinkedToTheElementsItDelimits(): void
    {
        $documentXml = '<w:p>
            <w:r><w:t xml:space="preserve">Before </w:t></w:r>
            <w:commentRangeStart w:id="0"/>
            <w:r><w:t>commented</w:t></w:r>
            <w:r><w:t xml:space="preserve"> text</w:t></w:r>
            <w:commentRangeEnd w:id="0"/>
            <w:r><w:commentReference w:id="0"/></w:r>
            <w:r><w:t xml:space="preserve"> after</w:t></w:r>
        </w:p>';
        $commentsXml = '<w:comment w:id="0" w:author="Author name" w:date="2024-01-01T00:00:00Z" w:initials="AN">
            <w:p><w:r><w:t>Comment text</w:t></w:r></w:p>
        </w:comment>';

        $phpWord = $this->getDocumentFromString(['document' => $documentXml, 'comments' => $commentsXml]);

        self::assertEquals(1, $phpWord->getComments()->countItems());
        /** @var Comment $comment */
        $comment = $phpWord->getComments()->getItem(0);
        /** @var Text $commentText */
        $commentText = $comment->getElement(0);
        self::assertEquals('Comment text', $commentText->getText());

        /** @var TextRun $textRun */
        $textRun = $phpWord->getSection(0)->getElement(0);
        self::assertInstanceOf(TextRun::class, $textRun);
        // The run holding the comment reference doesn't create any element
        self::assertCount(4, $textRun->getElements());

        /** @var Text $startElement */
        $startElement = $comment->getStartElement();
        self::assertInstanceOf(Text::class, $startElement);
        self::assertSame($textRun->getElement(1), $startElement);
        self::assertEquals('commented', $startElement->getText());
        /** @var Text $endElement */
        $endElement = $comment->getEndElement();
        self::assertInstanceOf(Text::class, $endElement);
        self::assertSame($textRun->getElement(2), $endElement);
        self::assertEquals(' text', $endElement->getText());
    }

    /**
     * A comment range can begin in a paragraph and end in another one.
     */
    public function testCommentRangeSpanningSeveralParagraphs(): void
    {
        $documentXml = '<w:p>
            <w:commentRangeStart w:id="1"/>
            <w:r><w:t>First</w:t></w:r>
        </w:p>
        <w:p>
            <w:r><w:t>Second</w:t></w:r>
            <w:commentRangeEnd w:id="1"/>
            <w:r><w:commentReference w:id="1"/></w:r>
        </w:p>';
        $commentsXml = '<w:comment w:id="1" w:author="Author name" w:date="2024-01-01T00:00:00Z" w:initials="AN">
            <w:p><w:r><w:t>Comment text</w:t></w:r></w:p>
        </w:comment>';

        $phpWord = $this->getDocumentFromString(['document' => $documentXml, 'comments' => $commentsXml]);

        /** @var Comment $comment */
        $comment = $phpWord->getComments()->getItem(0);
        /** @var TextRun $firstParagraph */
        $firstParagraph = $phpWord->getSection(0)->getElement(0);
        /** @var TextRun $secondParagraph */
        $secondParagraph = $phpWord->getSection(0)->getElement(1);
        self::assertSame($firstParagraph->getElement(0), $comment->getStartElement());
        self::assertSame($secondParagraph->getElement(0), $comment->getEndElement());
    }

    /**
     * Several comments can share the same beginning and the same end.
     */
    public function testNestedCommentRanges(): void
    {
        $documentXml = '<w:p>
            <w:commentRangeStart w:id="0"/>
            <w:commentRangeStart w:id="1"/>
            <w:r><w:t>First</w:t></w:r>
            <w:r><w:t>Second</w:t></w:r>
            <w:commentRangeEnd w:id="1"/>
            <w:commentRangeEnd w:id="0"/>
            <w:r><w:commentReference w:id="0"/></w:r>
            <w:r><w:commentReference w:id="1"/></w:r>
        </w:p>';
        $commentsXml = '<w:comment w:id="0" w:author="Author name" w:date="2024-01-01T00:00:00Z" w:initials="AN">
            <w:p><w:r><w:t>Outer comment</w:t></w:r></w:p>
        </w:comment>
        <w:comment w:id="1" w:author="Author name" w:date="2024-01-01T00:00:00Z" w:initials="AN">
            <w:p><w:r><w:t>Inner comment</w:t></w:r></w:p>
        </w:comment>';

        $phpWord = $this->getDocumentFromString(['document' => $documentXml, 'comments' => $commentsXml]);

        self::assertEquals(2, $phpWord->getComments()->countItems());
        /** @var Comment $outerComment */
        $outerComment = $phpWord->getComments()->getItem(0);
        /** @var Comment $innerComment */
        $innerComment = $phpWord->getComments()->getItem(1);
        /** @var TextRun $textRun */
        $textRun = $phpWord->getSection(0)->getElement(0);

        self::assertSame($textRun->getElement(0), $outerComment->getStartElement());
        self::assertSame($textRun->getElement(1), $outerComment->getEndElement());
        self::assertSame($textRun->getElement(0), $innerComment->getStartElement());
        self::assertSame($textRun->getElement(1), $innerComment->getEndElement());
    }

    /**
     * A comment which is only delimited by a range, without any reference, is
     * still linked to the document.
     */
    public function testCommentRangeWithoutReference(): void
    {
        $documentXml = '<w:p>
            <w:r><w:t>Before </w:t></w:r>
            <w:commentRangeStart w:id="0"/>
            <w:r><w:t>commented</w:t></w:r>
            <w:commentRangeEnd w:id="0"/>
        </w:p>';
        $commentsXml = '<w:comment w:id="0" w:author="Author name" w:date="2024-01-01T00:00:00Z" w:initials="AN">
            <w:p><w:r><w:t>Comment text</w:t></w:r></w:p>
        </w:comment>';

        $phpWord = $this->getDocumentFromString(['document' => $documentXml, 'comments' => $commentsXml]);

        /** @var Comment $comment */
        $comment = $phpWord->getComments()->getItem(0);
        /** @var TextRun $textRun */
        $textRun = $phpWord->getSection(0)->getElement(0);
        self::assertSame($textRun->getElement(1), $comment->getStartElement());
        self::assertSame($textRun->getElement(1), $comment->getEndElement());
    }

    /**
     * A comment without any range is linked to the element holding its reference.
     */
    public function testCommentWithoutRange(): void
    {
        $documentXml = '<w:p>
            <w:r><w:t>Text</w:t></w:r>
            <w:r><w:commentReference w:id="0"/></w:r>
        </w:p>';
        $commentsXml = '<w:comment w:id="0" w:author="Author name" w:date="2024-01-01T00:00:00Z" w:initials="AN">
            <w:p><w:r><w:t>Comment text</w:t></w:r></w:p>
        </w:comment>';

        $phpWord = $this->getDocumentFromString(['document' => $documentXml, 'comments' => $commentsXml]);

        /** @var Comment $comment */
        $comment = $phpWord->getComments()->getItem(0);
        /** @var TextRun $textRun */
        $textRun = $phpWord->getSection(0)->getElement(0);
        self::assertSame($textRun->getElement(0), $comment->getStartElement());
        self::assertSame($textRun->getElement(0), $comment->getEndElement());
    }
}
