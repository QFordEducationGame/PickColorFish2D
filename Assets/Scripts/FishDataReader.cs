using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using System.Data;
using Chart3D.DataProvider.Data.Excel;
using System.IO;
/// <summary>
/// Parse fish data from an excel file.
/// Use NOPI Plugin
/// 走配置是为了扩展性更好、因资源是非自动生成，使用配置比较容易实现实现颜色比对
/// </summary>
public static class FishDataReader 
{
     static readonly DataTable dt;
     static List<Fish> fishList;
    /// <summary>
    /// Load excel data
    /// </summary>
     static FishDataReader()
    {
        string path = Path.Combine(Application.streamingAssetsPath, "PickColorFishData.xlsx");
        dt = ExcelUtils.ReadAsDataTable(path);
        if (dt!=null && dt.Rows.Count > 0)
        {
            fishList = new List<Fish>(dt.Rows.Count);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string fishID = dt.Rows[i]["FishID"].ToString();
                int shape = int.Parse( dt.Rows[i]["Shape"].ToString());
                int red = int.Parse(dt.Rows[i]["Red"].ToString());
                int green = int.Parse(dt.Rows[i]["Green"].ToString());
                int yellow = int.Parse(dt.Rows[i]["Yellow"].ToString());
                int blue = int.Parse(dt.Rows[i]["Blue"].ToString());
                int white = int.Parse(dt.Rows[i]["White"].ToString());
                int purple = int.Parse(dt.Rows[i]["Purple"].ToString());
                Fish fish = new Fish(fishID, shape, red, green, yellow, blue, white, purple);
                fishList.Add(fish);
            }
            //Debug.Log("FishDataReader.cs:fishList's count = "+fishList.Count);
        }
    }
    /// <summary>
    /// for test data output
    /// </summary>
    public static void LogData()
    {
        foreach(var fish in fishList)
        {
            Debug.Log(string.Format("{0}:({1},{2},{3},{4},{5},{6})",
                fish.FishID, fish.Red, fish.Green, fish.Yellow, fish.Blue, fish.White, fish.Purple));
        }

    }

    public static List<Fish> GetFishList()
    {
        return fishList;
    }
}

public class Fish
{ 
    public string FishID
    { get; set; }

    public int Shape
    { get; set; }

    public int Red
    { get; set; }

    public int Green
    { get; set; }

    public int Yellow
    { get; set; }

    public int Blue
    { get; set; }

    public int White
    { get; set; }

    public int Purple
    { get; set; }

    public Fish(string fishID,int shape,int red,int green,int yellow,int blue,int white,int purple)
    {
        this.FishID = fishID;
        this.Shape = shape;
        this.Red = red;
        this.Green = green;
        this.Yellow = yellow;
        this.Blue = blue;
        this.White = white;
        this.Purple = purple;
    }

}