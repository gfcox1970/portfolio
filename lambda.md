# How to Create a Two-Column Table with All Possible Combinations of Two Ranges in Excel

If you have ever tried to create a two-column table that contains all possible combinations of two ranges of values in Excel, you might have encountered a variety of struggles. You may have used nested loops, helper columns, array formulas, or even VBA macros to achieve this task. However, these methods can be cumbersome, inefficient, or error-prone, especially if you have large or dynamic ranges.

Fortunately, there is a better way to create such a table with a single formula, using some of the new features introduced in Excel 365. In this blog post, I will show you how to use the `LET`, `UNIQUE`, `ROWS`, `SEQUENCE`, `MOD`, `SORT`, `INDEX`, and `HSTACK` functions to create a two-column table with all possible combinations of two ranges in Excel. I will also explain how to convert this formula into a reusable LAMBDA function, which can simplify your work even further.

## What is the MOD Function and How to Use It

Before we dive into the formula, let's take a moment to understand one of the key functions that we will use: the `MOD` function. The `MOD` function returns the remainder after dividing one number by another. For example, `=MOD(7, 3)` returns 1, because 7 divided by 3 gives a quotient of 2 and a remainder of 1.

The `MOD` function can be useful for various purposes, such as:

- Checking if a number is divisible by another number. For example, `=MOD(12, 4) = 0` returns TRUE, because 12 is divisible by 4.
- Wrapping around a sequence of numbers. For example, `=MOD(ROW()-1, 5) + 1` returns a sequence of numbers from 1 to 5, repeated indefinitely.
- Creating cyclical patterns or alternating values. For example, `=MOD(COLUMN(), 2) * 10` returns a pattern of 0 and 10, depending on the column number.

In our formula, we will use the `MOD` function to create sequences of numbers that will help us index the values from the two ranges.

## How to Create a Two-Column Table with All Possible Combinations of Two Ranges in Excel

Suppose we have two ranges of values in Excel, such as:

| A | B |
| - | - |
| Red | Apple |
| Green | Banana |
| Blue | Cherry |

We want to create a two-column table that contains all possible combinations of these values, such as:

| Col 1 | Col 2 |
| - | - |
| Red | Apple |
| Green | Apple |
| Blue | Apple |
| Red | Banana |
| Green | Banana |
| Blue | Banana |
| Red | Cherry |
| Green | Cherry |
| Blue | Cherry |

To do this, we can use the following formula:

```
=LET(
    a, A2:A15,
    b, B2:B15,
    a_count, ROWS(UNIQUE(a)),
    b_count, ROWS(UNIQUE(b)),
    total_rows, a_count * b_count,
    col_one_seq, SORT(MOD(SEQUENCE(total_rows, , 0), a_count) + 1),
    col_one_labels, INDEX(a, col_one_seq),
    col_two_seq, MOD(SEQUENCE(total_rows, , 0), b_count) + 1,
    col_two_labels, INDEX(b, col_two_seq),
    HSTACK(col_one_labels, col_two_labels)
)
```

Let's break down this formula and see how it works.

First, we use the `LET` function to define some variables that will make our formula easier to read and maintain. We assign the names `a` and `b` to the two ranges of values that we want to combine. We also calculate the number of unique values in each range, using the `ROWS` and `UNIQUE` functions, and assign them to the names `a_count` and `b_count`. Finally, we calculate the total number of rows that our table will have, by multiplying `a_count` and `b_count`, and assign it to the name `total_rows`.

Next, we create a sequence of numbers that will help us index the values from the first range (`a`). We use the `SEQUENCE` function to generate a sequence of numbers from 0 to `total_rows - 1`. Then, we use the `MOD` function to wrap around this sequence by `a_count`, and add 1 to get a sequence of numbers from 1 to `a_count`, repeated `b_count` times. For example, if `a_count` is 3 and `b_count` is 2, the sequence will be 1, 2, 3, 1, 2, 3. We use the `SORT` function to sort this sequence in ascending order, so that the values from the first range (`a`) will appear in order. We assign this sequence to the name `col_one_seq`.

Then, we use the `INDEX` function to return the values from the first range (`a`), indexed by `col_one_seq`. This will create a range of cells that contains the values from `a`, repeated and sorted according to `col_one_seq`. We assign this range to the name `col_one_labels`.

Similarly, we create another sequence of numbers that will help us index the values from the second range (`b`). We use the same `SEQUENCE` and `MOD` functions as before, but this time we wrap around the sequence by `b_count`, and add 1 to get a sequence of numbers from 1 to `b_count`, repeated `a_count` times. For example, if `a_count` is 3 and `b_count` is 2, the sequence will be 1, 1, 1, 2, 2, 2. We assign this sequence to the name `col_two_seq`.

Then, we use the `INDEX` function to return the values from the second range (`b`), indexed by `col_two_seq`. This will create a range of cells that contains the values from `b`, repeated and wrapped around according to `col_two_seq`. We assign this range to the name `col_two_labels`.

Finally, we use the `HSTACK` function to stack the two ranges horizontally, creating a two-column table that contains all possible combinations of `a` and `b`. We return this table as the result of the `LET` function.

## How to Convert the Formula into a LAMBDA Function

If you want to reuse this formula for different ranges of values, you can convert it into a LAMBDA function, which is a new feature in Excel 365 that allows you to create your own custom functions. To do this, you can follow these steps:

- Go to the Formulas tab, and click on Name Manager.
- Click on New, and enter a name for your function, such as `COMBINE_RANGES`.
- In the Refers to box, enter the formula that we created above, but replace the references to `a` and `b` with the parameters `_a` and `_b`. For example, `=LET(_a, _a, _b, _b, ...)`.
- Click on OK, and close the Name Manager.
- To use the function, enter `=COMBINE_RANGES(range1, range2)` in any cell, where `range1` and `range2` are the two ranges of values that you want to combine. Columns from a table can also be referenced using the syntax of `TableName[NameOfColumn]`

```
    =LAMBDA(
        _a, _b,
            LET(
                a, _a,
                b, _b,
                a_count, ROWS(UNIQUE(a)),
                b_count, ROWS(UNIQUE(b)),
                total_rows, a_count * b_count,
                col_one_seq, SORT(MOD(SEQUENCE(total_rows, , 0), a_count) + 1),
                col_one_labels, INDEX(a, col_one_seq),
                col_two_seq, MOD(SEQUENCE(total_rows, , 0), b_count) + 1,
                col_two_labels, INDEX(b, col_two_seq),
                HSTACK(col_one_labels, col_two_labels)
            ),
    )
```

## Conclusion

In this blog post, we learned how to use the `LET`, `UNIQUE`, `ROWS`, `SEQUENCE`, `MOD`, `SORT`, `INDEX`, and `HSTACK` functions to create a two-column table with all possible combinations of two ranges in Excel. We also learned how to convert this formula into a LAMBDA function, which can make it easier to reuse and maintain. We hope you found this post useful and interesting. 