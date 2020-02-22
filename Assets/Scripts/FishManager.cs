using System.Collections;
using System.Collections.Generic;
using System.Security.Cryptography;
using UnityEngine;
using UnityEngine.AddressableAssets;
using UnityEngine.ResourceManagement.AsyncOperations;
using System.Linq;

public class FishManager : MonoBehaviour
{
    public List<ColorFish> FishList;
    /// <summary>
    /// why set 41 failed?
    /// </summary>
    static Sprite[] spriteArrayFishAll = new Sprite[42];

    const int FishSpriteAltasCount = 6;
    int FishSpriteAltasLoadedCount = 0;
    int FishSpriteEachLoadedCount = 0;

    List<Fish> fishListData;

    public static Sprite GetFishSprite(string fishID)
    {
       Sprite sprite = spriteArrayFishAll.SingleOrDefault(s => s.name == fishID);
        if (fishID == "Fish1_7")
            System.Diagnostics.Debugger.Break();
        return sprite;
        
    }

    void Start()
    {
        //FishDataReader.LogData();
        LoadFish();
    }
    private void OnServerInitialized()
    {
        
    }
    /// <summary>
    /// load assets into a Scene with a filename and without the drawbacks of the Resources folder.
    /// https://gamedevbeginner.com/how-to-change-a-sprite-from-a-script-in-unity-with-examples/
    /// </summary>
    void LoadFish()
    {
        // Loading fishData from an excel file.
        fishListData = FishDataReader.GetFishList();
        //fishList.Shuffle();
        fishListData.ShuffleBetter();

        /*
        foreach (var fish in fishListData)
        {
            Debug.Log(string.Format("{0}:({1},{2},{3},{4},{5},{6})",
                fish.FishID, fish.Red, fish.Green, fish.Yellow, fish.Blue, fish.White, fish.Purple));
        }
        */

        //Loading fish sprite altas
        AsyncOperationHandle<Sprite[]> spriteHandle =
            Addressables.LoadAssetAsync<Sprite[]>("Assets/GameResources/Fishes/Fish1_1.png");
        AsyncOperationHandle<Sprite[]> spriteHandle1 =
            Addressables.LoadAssetAsync<Sprite[]>("Assets/GameResources/Fishes/Fish1_2.png");
        AsyncOperationHandle<Sprite[]> spriteHandle2 =
            Addressables.LoadAssetAsync<Sprite[]>("Assets/GameResources/Fishes/Fish1_3.png");
        AsyncOperationHandle<Sprite[]> spriteHandle3 =
            Addressables.LoadAssetAsync<Sprite[]>("Assets/GameResources/Fishes/Fish2.png");
        AsyncOperationHandle<Sprite[]> spriteHandle4 =
            Addressables.LoadAssetAsync<Sprite[]>("Assets/GameResources/Fishes/Fish3.png");
        AsyncOperationHandle<Sprite[]> spriteHandle5 =
            Addressables.LoadAssetAsync<Sprite[]>("Assets/GameResources/Fishes/Fish4.png");

        spriteHandle.Completed += LoadSpritesWhenReady;
        spriteHandle1.Completed += LoadSpritesWhenReady;
        spriteHandle2.Completed += LoadSpritesWhenReady;
        spriteHandle3.Completed += LoadSpritesWhenReady;
        spriteHandle4.Completed += LoadSpritesWhenReady;
        spriteHandle5.Completed += LoadSpritesWhenReady;
    }

    void LoadSpritesWhenReady(AsyncOperationHandle<Sprite[]> handleToCheck)
    {
        if (handleToCheck.Status == AsyncOperationStatus.Succeeded)
        {
            FishSpriteAltasLoadedCount++;

            for (int i = 0; i<handleToCheck.Result.Length;i++)
            {
                //Debug.Log("FishManager.cs:FishSpriteEachLoadedCount=" + FishSpriteEachLoadedCount);
                spriteArrayFishAll[FishSpriteEachLoadedCount] = handleToCheck.Result[i];
                FishSpriteEachLoadedCount++;
            }

            if (FishSpriteAltasLoadedCount == FishSpriteAltasCount)
            {
                //Debug.Log("FishManager.cs: LoadAllSprite Success,count=" +
                    //FishSpriteAltasCount);
                SetFishListData();
            }
        }
        else
        {
            Debug.Log("FishManager.cs: LoadSpriteAltasFailed!");
        }
    }

    void SetFishListData()
    {
        int i = 0;
        for (int i1 = 0; i1 < FishList.Count; i1++)
        {
            ColorFish fish = FishList[i1];
            if (fish)
                fish.SetFishData(fishListData[i++]);
            else
                Debug.LogError("FishManager.cs FishList null element!");

        }
    }
}
