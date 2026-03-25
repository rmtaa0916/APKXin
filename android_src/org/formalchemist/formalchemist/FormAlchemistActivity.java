package org.formalchemist.formalchemist;

import android.graphics.Bitmap;
import android.graphics.BitmapFactory;
import android.graphics.Color;
import android.graphics.drawable.BitmapDrawable;
import android.os.Bundle;
import android.os.Handler;
import android.os.Looper;
import android.view.Gravity;
import android.view.View;
import android.view.ViewGroup;
import android.widget.FrameLayout;
import android.widget.ImageView;

import org.kivy.android.PythonActivity;

import java.io.InputStream;

public class FormAlchemistActivity extends PythonActivity {
    private View splashOverlay;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);

        final ViewGroup content = findViewById(android.R.id.content);
        if (content == null) {
            return;
        }

        FrameLayout overlay = new FrameLayout(this);
        overlay.setLayoutParams(new FrameLayout.LayoutParams(
                ViewGroup.LayoutParams.MATCH_PARENT,
                ViewGroup.LayoutParams.MATCH_PARENT
        ));
        overlay.setBackgroundColor(Color.BLACK);

        ImageView splash = new ImageView(this);
        splash.setLayoutParams(new FrameLayout.LayoutParams(
                ViewGroup.LayoutParams.MATCH_PARENT,
                ViewGroup.LayoutParams.MATCH_PARENT,
                Gravity.CENTER
        ));
        splash.setScaleType(ImageView.ScaleType.CENTER_CROP);

        try {
            InputStream is = getAssets().open("presplash.png");
            Bitmap bmp = BitmapFactory.decodeStream(is);
            is.close();
            if (bmp != null) {
                splash.setImageDrawable(new BitmapDrawable(getResources(), bmp));
            }
        } catch (Exception ignored) {
            // If the asset cannot be opened, the black background remains and app still starts.
        }

        overlay.addView(splash);
        content.addView(overlay);
        splashOverlay = overlay;

        new Handler(Looper.getMainLooper()).postDelayed(new Runnable() {
            @Override
            public void run() {
                if (splashOverlay == null) {
                    return;
                }
                splashOverlay.animate()
                        .alpha(0f)
                        .setDuration(1800)
                        .withEndAction(new Runnable() {
                            @Override
                            public void run() {
                                try {
                                    content.removeView(splashOverlay);
                                } catch (Exception ignored) {
                                }
                                splashOverlay = null;
                            }
                        })
                        .start();
            }
        }, 900);
    }
}
