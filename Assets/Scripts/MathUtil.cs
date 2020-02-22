using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using UnityEngine;


public static class MathUtil
{
    /// <summary>
    /// 打乱数组中的元素
    /// </summary>
    /// <param name="intArray"></param>
    /// <returns></returns>
    public static int[] ShuffleArray( int[] intArray)
    {
        int[] newArray = intArray.Clone() as int[];
        for(int i=0;i<newArray.Length;i++)
        {
            int temp = newArray[i];
            int r = UnityEngine.Random.Range(i, newArray.Length);
            newArray[i] = newArray[i];
            newArray[r] = temp;
        }
        return newArray;
    }
    /// <summary>
    /// https://stackoverflow.com/questions/273313/randomize-a-listt
    /// </summary>
    private static System.Random rng = new System.Random();

    public static void Shuffle<T>(this IList<T> list)
    {
        int n = list.Count;
        while (n > 1)
        {
            n--;
            int k = rng.Next(n + 1);
            T value = list[k];
            list[k] = list[n];
            list[n] = value;
        }
    }

    public static void ShuffleBetter<T>(this IList<T> list)
    {
        RNGCryptoServiceProvider provider = new RNGCryptoServiceProvider();
        int n = list.Count;
        while (n > 1)
        {
            byte[] box = new byte[1];
            do provider.GetBytes(box);
            while (!(box[0] < n * (Byte.MaxValue / n)));
            int k = (box[0] % n);
            n--;
            T value = list[k];
            list[k] = list[n];
            list[n] = value;
        }
    }
}
