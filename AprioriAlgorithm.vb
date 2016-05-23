    Public Class AprioriAlgorithm
        Public ItemSet As List(Of Item) = Nothing
        Public SupportCount As Integer = 0
        Public Structure Item
            Public item As String
            Public SupportCount As Integer
        End Structure
        Public Structure Transaction
            Public ItemSet As List(Of Item)
            Public SupportCount As Integer
        End Structure
        Private mUniqueTransactions As List(Of Transaction)
        ReadOnly Property UniqueTransactions As List(Of Transaction)
            Get
                Return mUniqueTransactions
            End Get
        End Property
        Private mCandidateTransactions As List(Of Transaction)
        ''' <summary>
        ''' Given then original transactions Generate a list of All posible transactions
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        ReadOnly Property CandidateTransactions As List(Of Transaction)
            Get
                Return mCandidateTransactions
            End Get
        End Property
        Private mFrequentTransactions As List(Of Transaction)
        ''' <summary>
        ''' Given a list of transaction the transactions that are above the support count are
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        ReadOnly Property FrequentTransactions As List(Of Transaction)
            Get
                Return mFrequentTransactions
            End Get

        End Property
        Private mUniqueItems As List(Of Item) 'candidates
        ReadOnly Property UniqueItems As List(Of Item)
            Get
                Return mUniqueItems
            End Get

        End Property
        Private mFrequentItems As List(Of Item) 'L1
        ''' <summary>
        ''' Given a list of transactions which are the items that are above the support count
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        ReadOnly Property FrequentItems As List(Of Item)
            Get
                Return mFrequentItems
            End Get
        End Property
        Private MinimunSuportCount As Integer 'Minimun SupportCount
        Private OriginalTransactions As List(Of Transaction) 'Total transaction dataset
        Private NumberOfTransactions As Integer = 0 'Number of transactions
        Public Sub New(ByVal _OriginalTransactions As List(Of Transaction), ByVal _MinimunSuportCount As Integer)
            If _OriginalTransactions IsNot Nothing And _MinimunSuportCount > 0 Then
                'populate
                OriginalTransactions = _OriginalTransactions
                MinimunSuportCount = _MinimunSuportCount
                'Count
                NumberOfTransactions = _OriginalTransactions.Count
                '---Generate CandidateTransactions & FrequentTransactions & FrequentItems

                '1: FrequentItems
                'GetUniqueItems and SupportCounts (C1) CandidateItems
                mUniqueItems = GetUniqueItems(_OriginalTransactions)
                'Frequent Items L1 generated
                mFrequentItems = GetFrequentItems(MinimunSuportCount, UniqueItems)

                '2:CandidateTransactions
                'Generate All Possible Transactions -
                mCandidateTransactions = GenerateCandidateTransactions(UniqueItems)

                '--------------------------------------------------------------------------
                '':NOTE: Need to get the candidates supportCount from UniqueTransactions
                ''if not in UniqueTransactions then Supportcount is 0
                '--------------------------------------------------------------------------

                '3:FrequentTransactions
                'GetUniqueItemSets and supportCounts
                mUniqueTransactions = GetTransactionSetSupportCounts(_OriginalTransactions)
                'Frequent itemsets K1 calculated
                mFrequentTransactions = GetFrequentItemsets(MinimunSuportCount, UniqueTransactions)
            Else
            End If

        End Sub
        Public Function GenerateCandidateTransactions(ByVal _Itemset As List(Of Item)) As List(Of Transaction)
            GenerateCandidateTransactions = Nothing
            Dim ItemArr(mFrequentItems.Count) As Item
            Dim Candidate As New Transaction
            For Each Item As Item In _Itemset
                Dim counter As Integer = 0
                For counter = 0 To UBound(ItemArr)
                    ItemArr(counter).item = Item.item
                    ItemArr(counter).SupportCount = Item.item
                    Candidate.ItemSet = CreateItemset(ItemArr)
                    GenerateCandidateTransactions.Add(Candidate)
                Next

            Next
        End Function
        Public Function CreateItemset(ByVal _Item() As Item) As List(Of Item)
            CreateItemset = Nothing
            For Each ChkItem As Item In _Item
                CreateItemset.Add(ChkItem)
            Next
        End Function
        Public Function GetFrequentItemsets(ByVal SupportCount As Integer,
                                                ByVal Itemset As List(Of Transaction)) As List(Of Transaction)
            GetFrequentItemsets = Nothing
            For Each ChkItemset As Transaction In Itemset
                If ChkItemset.SupportCount > MinimunSuportCount Then
                    GetFrequentItemsets.Add(ChkItemset)
                End If
            Next
        End Function
        Public Function GetFrequentItems(ByVal SupportCount As Integer, ByVal ItemList As List(Of Item)) As List(Of Item)
            GetFrequentItems = Nothing
            For Each ChkItem In ItemList
                If ChkItem.SupportCount >= SupportCount Then
                    GetFrequentItems.Add(ChkItem)
                End If
            Next
        End Function
        Public Function GetTransactionSetSupportCounts(ByVal _CandidateTransactions As List(Of Transaction)) As List(Of Transaction)
            'get frequencys of items
            Dim groups = _CandidateTransactions.GroupBy(Function(value) value)
            Dim UniqueLst As New List(Of Transaction)
            Dim NewUnique As New Transaction
            For Each grp In groups
                'save grouped items (distinct) and frequency (supportcount)
                NewUnique.ItemSet = grp(0).ItemSet
                NewUnique.SupportCount = grp.Count
                UniqueLst.Add(NewUnique)
            Next
            Return UniqueLst
        End Function
        Public Function GetFrequencys(ByVal _Item As Item) As Item
            Return _Item
        End Function
        Public Function GetFrequencys(ByVal _Itemset As List(Of Item)) As List(Of Item)
            Return _Itemset
        End Function
        ''' <summary>
        ''' given a set of transactions return a list of unique items and support counts
        ''' </summary>
        ''' <param name="_Transactions"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetUniqueItems(ByVal _Transactions As List(Of Transaction)) As List(Of Item)
            Dim _AllItems As New List(Of Item)
            'count how many unique transactions.SingleTransacton.item are in TotalTransactions
            For Each Newtransaction As Transaction In _Transactions
                'Get transaction (list of items)
                For Each NewItem As Item In Newtransaction.ItemSet
                    'gets all items
                    _AllItems.Add(NewItem)
                Next
            Next
            'get frequencys of items
            Dim groups = _AllItems.GroupBy(Function(value) value)
            Dim UniqueLst As New List(Of Item)
            Dim NewUnique As New Item
            For Each grp In groups
                'save grouped items (distinct) and frequency (supportcount)
                NewUnique.item = grp(0).ToString
                NewUnique.SupportCount = grp.Count
                UniqueLst.Add(NewUnique)
            Next
            Return UniqueLst
        End Function
        Public Function GetCandidateTransactions(ByVal _Transactions As List(Of Transaction),
                                  ByVal _CandidateItems As List(Of Item)) As List(Of Transaction)
            Dim Found(_CandidateItems.Count) As Boolean
            Dim Detected As Boolean = False
            Dim _CandidateTransactions As New List(Of Transaction)
            ''Check Each transaction
            For Each T As Transaction In _Transactions
                For I As Integer = 0 To _CandidateItems.Count
                    'check if itemset contains item
                    If T.ItemSet.Contains(_CandidateItems(I)) Then
                        Found(I) = True
                    End If
                Next
                'Detect if All of the items have matches in itemset
                Dim Counter As Integer = 0
                For f As Integer = 0 To UBound(Found)
                    If f = False Then
                        Detected = False
                        Counter += 1
                    End If
                Next
                'If Counter is Greater then 0 then a false has occured
                If Counter > 0 Then Detected = False
                'If True then Add to Matched candidates
                If Detected = True Then
                    _CandidateTransactions.Add(T)
                End If
            Next
            Return _CandidateTransactions
        End Function
    End Class
