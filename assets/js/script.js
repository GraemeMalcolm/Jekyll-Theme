$(function() {
    $('article img').each(function() {
        var src = $(this).attr('src');
        $(this).wrap('<a href="' + src + '" target="_blank"></a>');
    });
    $('article > h2').each(function() {
        $('nav.toc > ul').append(
            $('<li>')
                .attr('class', 'nav-item')
                .append(
                    $('<a>')
                        .attr('class', 'nav-link')
                        .text($(this).text())
                        .attr('href', '#' + $(this).attr('id'))
                ) 
        );
    });
    $('[data-spy="scroll"]').each(function () {
        var $spy = $(this).scrollspy('refresh')
    });

    function setupMiniAnton() {
        var elements = {
            form: document.getElementById('mini-anton-form'),
            input: document.getElementById('mini-anton-input'),
            send: document.getElementById('mini-anton-send'),
            messages: document.getElementById('mini-anton-messages'),
            status: document.getElementById('mini-anton-status')
        };

        if (!elements.form || !elements.input || !elements.send || !elements.messages || !elements.status) {
            return;
        }

        var state = {
            indexData: null,
            keywordMap: new Map(),
            loadingIndex: false,
            indexReady: false,
            indexUrl: 'https://raw.githubusercontent.com/MicrosoftLearning/ai-apps/refs/heads/main/ask-anton/index.json'
        };

        function sanitizeHtml(text) {
            // Create a text node to safely escape HTML without double-encoding
            var div = document.createElement('div');
            div.textContent = text;
            return div.innerHTML;
        }

        function normalizeSearchText(text) {
            return text.toLowerCase().trim().replace(/[^a-z0-9\s]/g, ' ').replace(/\s+/g, ' ').trim();
        }

        function getSearchIntentQuery(text) {
            var trimmedText = text.trim();
            var lowerText = trimmedText.toLowerCase();

            if (lowerText.indexOf('search ') === 0) {
                return trimmedText.slice(7).trim();
            }

            if (lowerText.indexOf('find ') === 0) {
                return trimmedText.slice(5).trim();
            }

            if (lowerText.indexOf('documentation') >= 0 || lowerText.indexOf('docs') >= 0 || lowerText.indexOf('microsoft learn') >= 0) {
                return trimmedText;
            }

            return null;
        }

        function extractBingSearchKeywords(text) {
            var words = normalizeSearchText(text).split(' ').filter(Boolean);
            var stopWords = new Set([
                'a', 'an', 'and', 'are', 'as', 'at', 'be', 'by', 'for', 'from', 'in', 'is', 'it', 'its', 'of', 'on', 'that', 'the', 'to', 'with', 'or', 'but',
                'if', 'than', 'then', 'so', 'after', 'before', 'between', 'during', 'into', 'through', 'over', 'under', 'until', 'up', 'down', 'out', 'off',
                'above', 'below', 'i', 'you', 'he', 'she', 'we', 'they', 'me', 'him', 'her', 'us', 'them', 'my', 'your', 'his', 'our', 'their', 'this',
                'these', 'those', 'some', 'any', 'all', 'each', 'every', 'both', 'few', 'more', 'most', 'such', 'no', 'nor', 'not', 'only', 'own', 'same',
                'other', 'another', 'much', 'many', 'am', 'was', 'were', 'been', 'being', 'have', 'has', 'had', 'do', 'does', 'did', 'can', 'could', 'would',
                'should', 'may', 'might', 'must', 'shall', 'will', 'get', 'make', 'know', 'see', 'take', 'come', 'go', 'want', 'use', 'find', 'need', 'try',
                'ask', 'work', 'help', 'like', 'seem', 'become', 'let', 'tell', 'show', 'give', 'provide', 'explain', 'describe', 'define', 'what', 'when',
                'where', 'who', 'how', 'why', 'which', 'whom', 'whose', 'also', 'just', 'now', 'here', 'there', 'very', 'too', 'really', 'still', 'always',
                'never', 'often', 'sometimes', 'maybe', 'perhaps', 'about', 'yes', 'thing', 'something', 'anything', 'nothing', 'everything', 'someone',
                'anyone', 'everyone', 'understand', 'think', 'believe', 'feel', 'appear', 'say', 'anton', 'please', 'using', 'search', 'documentation',
                'docs', 'learn', 'details', 'overview'
            ]);
            var uniqueWords = [];
            var seenWords = new Set();

            words.forEach(function(word) {
                if (word.length < 2 || stopWords.has(word) || seenWords.has(word)) {
                    return;
                }
                seenWords.add(word);
                uniqueWords.push(word);
            });

            return uniqueWords.join(' ');
        }

        function formatMessageText(text) {
            // First sanitize the text, then replace line breaks with br tags
            var sanitized = sanitizeHtml(text);
            return sanitized.replace(/\n\n/g, '<br><br>').replace(/\n/g, '<br>');
        }

        function typeMessage(element, htmlContent, speed) {
            var tempDiv = document.createElement('div');
            tempDiv.innerHTML = htmlContent;
            
            var textNodes = [];
            var walk = document.createTreeWalker(
                tempDiv,
                NodeFilter.SHOW_TEXT,
                null,
                false
            );
            
            var node;
            while (node = walk.nextNode()) {
                if (node.textContent.trim()) {
                    textNodes.push(node);
                }
            }
            
            var charIndex = 0;
            var nodeIndex = 0;
            var totalChars = 0;
            
            textNodes.forEach(function(n) {
                totalChars += n.textContent.length;
            });
            
            function type() {
                if (nodeIndex >= textNodes.length) {
                    return;
                }
                
                var currentNode = textNodes[nodeIndex];
                var currentText = currentNode.textContent;
                
                if (charIndex < currentText.length) {
                    currentNode.textContent = currentText.substring(0, charIndex + 1);
                    charIndex++;
                    setTimeout(type, speed);
                } else {
                    nodeIndex++;
                    charIndex = 0;
                    setTimeout(type, speed);
                }
            }
            
            element.innerHTML = htmlContent.replace(/>[^<]*</g, function(match) {
                return match.substring(0, 1) + match.substring(match.length - 1);
            });
            
            var display = element.cloneNode(true);
            display.innerHTML = '';
            for (var i = 0; i < tempDiv.children.length; i++) {
                display.appendChild(tempDiv.children[i].cloneNode(true));
            }
            
            element.innerHTML = htmlContent;
            var allText = element.innerText;
            element.innerHTML = '<p></p>';
            var pElement = element.querySelector('p');
            var charIdx = 0;
            
            function typeChar() {
                if (charIdx < allText.length) {
                    var char = allText[charIdx];
                    pElement.textContent += char;
                    charIdx++;
                    elements.messages.scrollTop = elements.messages.scrollHeight;
                    setTimeout(typeChar, speed);
                }
            }
            
            typeChar();
        }

        function addMessage(role, content, isHtml, animate) {
            
            // For user messages, always use textContent to prevent HTML injection
            // For assistant messages, treat isHtml flag but sanitize for safety
            var messageText = isHtml ? content : content;
            var messageHtml = isHtml ? content : formatMessageText(content);
            
            if (animate && role === 'assistant') {
                var pElement = document.createElement('p');
                messageDiv.appendChild(pElement);
                elements.messages.appendChild(messageDiv);
                elements.messages.scrollTop = elements.messages.scrollHeight;
                
                // For animated messages, set innerHTML safely and extract text for typing effect
                if (isHtml) {
                    pElement.innerHTML = messageHtml;
                } else {
                    pElement.innerHTML = messageHtml;
                }
                
                var allText = pElement.textContent; // Get rendered text content
                pElement.textContent = ''; // Clear and rebuild with typing effect
                var charIdx = 0;
                
                function typeChar() {
                    if (charIdx < allText.length) {
                        var char = allText[charIdx];
                        pElement.textContent += char;
                        charIdx++;
                        elements.messages.scrollTop = elements.messages.scrollHeight;
                        setTimeout(typeChar, 15);
                    }
                }
                
                typeChar();
            } else {
                var pElement = document.createElement('p');
                if (isHtml && role === 'assistant') {
                    // Only use innerHTML for trusted assistant content
                    pElement.innerHTML = messageHtml;
                } else {
                    // For user messages, use textContent to prevent HTML injection
                    pElement.textContent = content;
                }
                messageDiv.appendChild(pElement)
            } else {
                messageDiv.innerHTML = '<p>' + messageHtml + '</p>';
                elements.messages.appendChild(messageDiv);
                elements.messages.scrollTop = elements.messages.scrollHeight;
            }
        }

        function buildKeywordMap() {
            state.keywordMap = new Map();
            state.indexData.forEach(function(category) {
                category.documents.forEach(function(doc) {
                    doc.keywords.forEach(function(keyword) {
                        var normalizedKeyword = keyword.toLowerCase().trim();
                        if (normalizedKeyword) {
                            state.keywordMap.set(normalizedKeyword, {
                                document: doc,
                                category: category.category,
                                link: category.link
                            });
                        }
                    });
                });
            });
        }

        function performSearch(question) {
            var normalizedQuestion = normalizeSearchText(question);
            var words = normalizedQuestion.split(' ').filter(Boolean);
            var nGrams = [];
            var i;
            var stopWords = ['what', 'is', 'are', 'the', 'a', 'an', 'how', 'does', 'do', 'can', 'about', 'tell', 'me', 'explain', 'describe', 'show', 'give', 'anton', 'i', 'you', 'he', 'she', 'it', 'we', 'they', 'my', 'your', 'his', 'her', 'its', 'our', 'their', 'why', 'which', 'whom', 'whose', 'all', 'any', 'this', 'that', 'these', 'those'];

            for (i = 0; i <= words.length - 3; i++) {
                nGrams.push({ text: words.slice(i, i + 3).join(' '), length: 3 });
            }
            for (i = 0; i <= words.length - 2; i++) {
                nGrams.push({ text: words.slice(i, i + 2).join(' '), length: 2 });
            }
            words.forEach(function(word) {
                if (word.length >= 2 && stopWords.indexOf(word) < 0) {
                    nGrams.push({ text: word, length: 1 });
                }
            });

            var matchedKeywords = new Set();
            var documentMatches = new Map();

            nGrams.forEach(function(ngram) {
                var match = state.keywordMap.get(ngram.text);
                if (!match) {
                    return;
                }
                matchedKeywords.add(ngram.text);
                var docId = match.document.id;
                if (!documentMatches.has(docId)) {
                    documentMatches.set(docId, {
                        document: match.document,
                        category: match.category,
                        link: match.link,
                        matchedKeywords: []
                    });
                }
                documentMatches.get(docId).matchedKeywords.push(ngram.text);
            });

            var filteredKeywords = new Set();
            Array.from(matchedKeywords)
                .sort(function(a, b) { return b.split(' ').length - a.split(' ').length; })
                .forEach(function(keyword) {
                    var isSubset = false;
                    filteredKeywords.forEach(function(existing) {
                        if (!isSubset && existing !== keyword && existing.indexOf(keyword) >= 0) {
                            isSubset = true;
                        }
                    });
                    if (!isSubset) {
                        filteredKeywords.add(keyword);
                    }
                });

            var rankedMatches = [];
            documentMatches.forEach(function(match) {
                var validKeywords = match.matchedKeywords.filter(function(kw) {
                    return filteredKeywords.has(kw);
                });
                if (validKeywords.length > 0) {
                    rankedMatches.push({
                        document: match.document,
                        category: match.category,
                        link: match.link,
                        matchedKeywords: validKeywords
                    });
                }
            });

            rankedMatches.sort(function(a, b) {
                var aScore = a.matchedKeywords.reduce(function(sum, kw) { return sum + kw.split(' ').length; }, 0);
                var bScore = b.matchedKeywords.reduce(function(sum, kw) { return sum + kw.split(' ').length; }, 0);
                return bScore - aScore;
            });

            var categories = [];
            var links = [];
            var documents = [];
            rankedMatches.forEach(function(match) {
                documents.push(match.document);
                if (categories.indexOf(match.category) < 0) {
                    categories.push(match.category);
                }
                if (links.indexOf(match.link) < 0) {
                    links.push(match.link);
                }
            });

            return {
                categories: categories,
                links: links,
                documents: documents
            };
        }

        function buildLearnSearchUrl(queryText) {
            var bingKeywords = extractBingSearchKeywords(queryText) || normalizeSearchText(queryText);
            return {
                keywords: bingKeywords,
                url: 'https://learn.microsoft.com/en-us/search/?terms=' + encodeURIComponent(bingKeywords) + '&category=Documentation'
            };
        }

        function setBusy(isBusy) {
            elements.send.disabled = isBusy;
            elements.input.disabled = isBusy;
        }

        async function loadIndex() {
            if (state.indexReady || state.loadingIndex) {
                return;
            }

            state.loadingIndex = true;
            elements.status.textContent = 'Loading knowledge base...';

            try {
                var response = await fetch(state.indexUrl, { cache: 'no-store' });
                if (!response.ok) {
                    throw new Error('Failed to load index.');
                }
                state.indexData = await response.json();
                buildKeywordMap();
                state.indexReady = true;
                elements.status.textContent = 'Ready.';
            } catch (error) {
                state.indexReady = false;
                elements.status.textContent = 'Unable to load index; search links still available.';
            } finally {
                state.loadingIndex = false;
            }
        }

        async function handleSubmit(event) {
            event.preventDefault();

            var userMessage = elements.input.value.trim();
            if (!userMessage) {
                return;
            }

            addMessage('user', userMessage, false);
            elements.input.value = '';
            elements.input.style.height = '30px';

            setBusy(true);
            await loadIndex();

            var searchQuery = getSearchIntentQuery(userMessage);
            if (searchQuery) {
                var searchTarget = buildLearnSearchUrl(searchQuery);
                elements.status.textContent = 'Created Microsoft Learn search link.';
                addMessage('assistant', 'I searched for "' + escapeHtml(searchTarget.keywords) + '". <a href="' + searchTarget.url + '" target="_blank" rel="noopener noreferrer">Here\'s what I found.</a>', true, true);
                setBusy(false);
                return;
            }

            if (!state.indexReady) {
                var fallbackSearch = buildLearnSearchUrl(userMessage);
                addMessage('assistant', 'I cannot use the knowledge base right now. You can <a href="' + fallbackSearch.url + '" target="_blank" rel="noopener noreferrer">search Microsoft Learn documentation</a>.', true, true);
                setBusy(false);
                return;
            }

            var searchResult = performSearch(userMessage);
            if (!searchResult.documents.length) {
                var noResultsSearch = buildLearnSearchUrl(userMessage);
                elements.status.textContent = 'No direct match in index.';
                addMessage('assistant', 'I do not have specific information about that topic. Try <a href="' + noResultsSearch.url + '" target="_blank" rel="noopener noreferrer">searching Microsoft Learn</a>.', true, true);
                setBusy(false);
                return;
            }

            elements.status.textContent = 'Found context in: ' + searchResult.categories.join(', ');
            var responseText = searchResult.documents.map(function(document) {
                return document.content;
            }).join('\n\n');
            addMessage('assistant', responseText, false, true);
            setBusy(false);
        }

        elements.form.addEventListener('submit', handleSubmit);
        elements.input.addEventListener('keydown', function(event) {
            if (event.key === 'Enter' && !event.shiftKey) {
                event.preventDefault();
                elements.form.requestSubmit();
            }
        });
        elements.input.addEventListener('input', function() {
            elements.input.style.height = 'auto';
            elements.input.style.height = Math.min(elements.input.scrollHeight, 72) + 'px';
        });

        loadIndex();
    }

    setupMiniAnton();

    $('pre').each(function(index) {
    var generatedId = 'codeBlock' + index;
    var languageClass = $(this).children('code:first').attr('class').split(' ')[0];
    var language = languageClass == 'language-sh' ? 'shell' : 
        languageClass == 'language-js' ? 'javascript' : 
        languageClass == 'language-xml' ? 'xml' : 
        languageClass == 'language-sql' ? 'sql' : 
        languageClass == 'language-csharp' ? 'c#' : 'code';
    $(this).attr('id', generatedId);
    var header = $('<div/>', {
      class: 'code-header mt-3 mb-0 bg-light d-flex justify-content-between border',
    }).append(
        $('<span/>', {     
            class: 'mx-2 text-muted text-capitalize font-weight-light',
        html: language
        })
    ).append(
        $('<button/>', {   
            class: 'm-0 btn btn-code btn-sm btn-light codeBtn rounded-0 border-left text-muted font-weight-light',
        'data-clipboard-target': '#' + generatedId,
        type: 'button'
        }).append(
        $('<i/>', {
            class: 'fa fa-files-o mr-2',
          'aria-hidden': 'true'
        })
      ).append(
        'Copy'
      )
    );
    header.insertBefore($(this));
    $(this).addClass('mt-0');
});
hljs.initHighlightingOnLoad();
new ClipboardJS('.btn-code');
});
