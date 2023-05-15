import java.util.Arrays;

public class DatasetNormalization {
    public static void main(String[] args) {
        int[][] dataset = {{10, 20, 30}, {55, 67, 90}, {100, 45, 98}};
        double[][] normalizedDataset = normalizeDataset(dataset);
        
        // Print normalized dataset
        for (double[] point : normalizedDataset) {
            System.out.println(Arrays.toString(point));
        }
    }
    
    public static double[][] normalizeDataset(int[][] dataset) {
        int numDimensions = dataset[0].length;
        
        // Find the maximum and minimum values for each dimension
        int[] maxVals = new int[numDimensions];
        int[] minVals = new int[numDimensions];
        
        Arrays.fill(maxVals, Integer.MIN_VALUE);
        Arrays.fill(minVals, Integer.MAX_VALUE);
        
        for (int[] point : dataset) {
            for (int i = 0; i < numDimensions; i++) {
                if (point[i] > maxVals[i]) {
                    maxVals[i] = point[i];
                }
                
                if (point[i] < minVals[i]) {
                    minVals[i] = point[i];
                }
            }
        }
        
        // Normalize the dataset
        double[][] normalizedDataset = new double[dataset.length][numDimensions];
        
        for (int i = 0; i < dataset.length; i++) {
            for (int j = 0; j < numDimensions; j++) {
                double normalizedValue = (dataset[i][j] - minVals[j]) / (double)(maxVals[j] - minVals[j]);
                normalizedDataset[i][j] = Math.round(normalizedValue * 10.0) / 10.0; // Round to one decimal place
            }
        }
        
        return normalizedDataset;
    }
}
