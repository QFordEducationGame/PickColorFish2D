using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEngine.AddressableAssets;
using UnityEngine.ResourceManagement.AsyncOperations;

public class ColorFish : MonoBehaviour
{
    Fish fishData = null;
    static Vector2 originalSize = new Vector2(120f, 280f);
    Vector3 originalScale;
 
    public void SetFishData(Fish fish)
    {
        originalScale = transform.localScale;
        fishData = fish;
        DrawFish();  
    }

    void DrawFish()
    {
        if (fishData == null)
        {
            Debug.Log("ColorFish.cs fishData null!" );
            return;
        }
        
        Sprite sprite = FishManager.GetFishSprite(fishData.FishID);
        if (sprite == null)
        {
            Debug.Log("ColorFish.cs Sprite not found! FishID = "+fishData.FishID);
            return;
        }

        if (fishData != null && sprite != null)
        {
            GetComponent<SpriteRenderer>().sprite = sprite;
            //修复因为切图大小不一致导致的sprite大小不一的问题
            float x = sprite.textureRect.width;
            float y = sprite.textureRect.height;
            float sFactor = (originalSize.x / x + originalSize.y / y) / 2;
            transform.localScale = originalScale * sFactor;
        }
    }
}
