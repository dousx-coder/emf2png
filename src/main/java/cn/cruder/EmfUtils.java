package cn.cruder;

import org.apache.poi.hemf.usermodel.HemfPicture;
import org.apache.poi.util.Units;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.geom.Dimension2D;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

/**
 * @Author: cruder
 * @Date: 2021/09/13/16:08
 */
public class EmfUtils {


    /**
     * 获取文件路径
     *
     * @param dirName  文件夹
     * @param fileName 文件名
     * @return
     * @throws IOException
     */
    private static String getFilePath(String dirName, String fileName) throws IOException {
        String resourcesPath = "src" + File.separator + "main" + File.separator + "resources";
        File directory = new File(resourcesPath);
        String courseFile = directory.getCanonicalPath();
        return courseFile + File.separator + dirName + File.separator + fileName;

    }



    public static void main(String[] args) throws IOException {

        String emfFilePath = getFilePath("emf", "81b7d45d-e25e-43fc-b7cb-79b437306f6d.emf");
        String pngFilePath = getFilePath("png", "81b7d45d-e25e-43fc-b7cb-79b437306f6d.png");
        emf2png(emfFilePath, pngFilePath);


    }

    /**
     * emf 转png
     * @param emfFilePath
     * @param pngFilePath
     * @throws IOException
     */
    private static void emf2png(String emfFilePath, String pngFilePath) throws IOException {
        File emfFile = new File(emfFilePath);
        try (FileInputStream fis = new FileInputStream(emfFile)) {
            // for EMF / EMF+
            HemfPicture emf = new HemfPicture(fis);
            Dimension2D dim = emf.getSize();
            int width = Units.pointsToPixel(dim.getWidth());
            // keep aspect ratio for height
            int height = Units.pointsToPixel(dim.getHeight());
            double max = Math.max(width, height);
            if (max > 1500) {
                width *= 1500/max;
                height *= 1500/max;
            }
            BufferedImage bufImg = new BufferedImage(width, height, BufferedImage.TYPE_INT_ARGB);
            Graphics2D g = bufImg.createGraphics();
            g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
            g.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
            g.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);
            g.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS, RenderingHints.VALUE_FRACTIONALMETRICS_ON);
            emf.draw(g, new Rectangle2D.Double(0,0,width,height));
            g.dispose();

            ImageIO.write(bufImg, "PNG",new File(pngFilePath));
        }
    }


}
